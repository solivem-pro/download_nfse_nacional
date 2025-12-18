import os
import base64
import gzip
import logging
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import tempfile
import time
import json
from pathlib import Path
from contextlib import contextmanager
from typing import Callable, Iterable, Optional
import xml.etree.ElementTree as ET
import requests

## Módulos auxiliares
from downloader.pdf import NFSePDFDownloader
from config.config import Config, STATUS_STOP, MAX_TENT

logger = logging.getLogger(__name__)

from cryptography.hazmat.primitives.serialization import (
    Encoding,
    PrivateFormat,
    NoEncryption,
)
from cryptography.hazmat.primitives.serialization.pkcs12 import (
    load_key_and_certificates,
)

## ------------------------------------------------------------------------------
## Classe de download por competência (dCompet)
## ------------------------------------------------------------------------------
class NFSeDownloaderCompetencia:
    """Baixa os arquivos NFSe por competência (campo dCompet)"""

    def __init__(self, config: Config):
        self.config = config
        self.logger = logging.getLogger(__name__)
        self.session: Optional[requests.Session] = None
        self.base_url = "https://adn.nfse.gov.br/contribuintes/DFe"
        self._running = True

    def stop(self):
        """Para a execução do download"""
        self._running = False
        if self.session:
            self.session.close()

    def running(self):
        """Verifica se o processo deve continuar"""
        return self._running

    @staticmethod
    def extrair_competencia(xml_bytes: bytes) -> tuple[str, str]:
        """Extrai ano e mês do campo dCompet (competência) do XML"""
        now = datetime.now()
        try:
            root = ET.fromstring(xml_bytes)
            
            # Procurar pelo campo dCompet em diferentes locais possíveis
            for tag in ("dCompet", "Competencia", "DataCompetencia"):
                el = root.find(f'.//{{*}}{tag}')
                if el is None:
                    el = root.find(f'.//*[local-name()="{tag}"]')
                
                if el is not None and el.text:
                    txt = el.text.strip()
                    try:
                        # Tentar formato ISO (AAAA-MM-DD)
                        dt = datetime.fromisoformat(txt.replace("Z", ""))
                        return str(dt.year), f"{dt.month:02d}"
                    except Exception:
                        # Tentar outros formatos
                        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%Y%m"):
                            try:
                                dt = datetime.strptime(txt[:10], fmt)
                                return str(dt.year), f"{dt.month:02d}"
                            except Exception:
                                continue
        except Exception as e:
            logger.debug(f"Erro ao extrair competência: {e}")
        
        # Se não encontrou, retorna data atual
        return str(now.year), f"{now.month:02d}"

    @staticmethod
    def extrair_data_emissao(xml_bytes: bytes) -> tuple[str, str]:
        """Extrai ano e mês da data de emissão (dhEmi, dhEvento, DataEmissao)"""
        now = datetime.now()
        try:
            root = ET.fromstring(xml_bytes)
            el = None
            for tag in ("dhEmi", "dhEvento", "DataEmissao"):
                el = root.find(f'.//{{*}}{tag}')
                if el is None:
                    el = root.find(f'.//*[local-name()="{tag}"]')
                
                if el is not None and el.text:
                    break
            
            if el is not None and el.text:
                txt = el.text.strip()
                try:
                    dt = datetime.fromisoformat(txt.replace("Z", ""))
                    return str(dt.year), f"{dt.month:02d}"
                except Exception:
                    for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
                        try:
                            dt = datetime.strptime(txt[:10], fmt)
                            return str(dt.year), f"{dt.month:02d}"
                        except Exception:
                            continue
        except Exception:
            pass
        
        return str(now.year), f"{now.month:02d}"

    def determinar_tipo_documento(self, xml_bytes: bytes) -> str:
        """Determina se o documento é PRESTADO, TOMADO ou EVENTO"""
        try:
            root = ET.fromstring(xml_bytes)
            
            # Tentar encontrar o CNPJ do prestador
            cnpj_prestador = None
            
            # Procurar em diferentes locais possíveis para o CNPJ do prestador
            for cnpj_tag in ['.//{*}CNPJ', './/{*}Cnpj', './/{*}prestador/{*}CNPJ', 
                            './/{*}Prestador/{*}Cnpj', './/{*}infDPS/{*}prest/{*}CNPJ']:
                element = root.find(cnpj_tag)
                if element is not None and element.text:
                    cnpj_prestador = element.text.strip()
                    break
            
            # Se encontrou CNPJ do prestador, comparar com CNPJ da empresa
            if cnpj_prestador:
                if cnpj_prestador == self.config.cnpj:
                    return "PRESTADOS"
                else:
                    return "TOMADOS"
            
            # Se não encontrou CNPJ do prestador, verificar se é evento
            evento_tags = ['{*}evento', '{*}Evento', '{*}InfEvento']
            for tag in evento_tags:
                if root.find(f'.//{tag}') is not None:
                    return "EVENTOS"
                    
        except Exception as e:
            self.logger.error(f"Erro ao determinar tipo do documento: {e}")
        
        # Default para eventos se não conseguir determinar
        return "EVENTOS"

    @contextmanager
    def pfx_to_pem(
        self,
        pfx_path: Optional[str] = None,
        pfx_password: Optional[str] = None,
    ) -> Iterable[str]:
        """Convert pfx_path to a temporary PEM file."""
        if pfx_path is None:
            pfx_path = self.config.cert_path
        if pfx_password is None:
            pfx_password = self.config.cert_pass
        data = Path(pfx_path).read_bytes()
        priv_key, cert, add_certs = load_key_and_certificates(
            data, pfx_password.encode(), None
        )
        tmp = tempfile.NamedTemporaryFile(suffix=".pem", delete=False)
        pem_path = tmp.name
        tmp.close()
        with open(pem_path, "wb") as f:
            f.write(priv_key.private_bytes(Encoding.PEM, PrivateFormat.PKCS8, NoEncryption()))
            f.write(cert.public_bytes(Encoding.PEM))
            if add_certs:
                for ca in add_certs:
                    f.write(ca.public_bytes(Encoding.PEM))
        try:
            yield pem_path
        finally:
            os.remove(pem_path)

    ## ------------------------------------------------------------------------------
    ## Tratamentos por execução
    ## ------------------------------------------------------------------------------
    def limpar_pastas_empresa(self):
        """Limpa o conteúdo das pastas PRESTADOS, TOMADOS e EVENTOS"""
        pastas = ["PRESTADOS", "TOMADOS", "EVENTOS"]
        
        for pasta in pastas:
            pasta_path = os.path.join(self.config.output_dir, pasta)
            if os.path.exists(pasta_path):
                for arquivo in os.listdir(pasta_path):
                    caminho_arquivo = os.path.join(pasta_path, arquivo)
                    try:
                        if os.path.isfile(caminho_arquivo):
                            os.remove(caminho_arquivo)
                        elif os.path.isdir(caminho_arquivo):
                            import shutil
                            shutil.rmtree(caminho_arquivo)
                    except Exception as e:
                        self.logger.error(f"Erro ao deletar {caminho_arquivo}: {e}")

    def criar_pastas_empresa(self):
        """Cria as pastas PRESTADOS, TOMADOS e EVENTOS se não existirem"""
        pastas = ["PRESTADOS", "TOMADOS", "EVENTOS"]
        
        for pasta in pastas:
            pasta_path = os.path.join(self.config.output_dir, pasta)
            os.makedirs(pasta_path, exist_ok=True)

    def registrar_erro(self, nsu: int, chave: str, tipo: str, descricao: str, 
                      ano_compet: str = None, mes_compet: str = None):
        """Registra erro no arquivo erros.txt da empresa"""
        try:
            if ano_compet is None or mes_compet is None:
                ano_compet = str(datetime.now().year)
                mes_compet = f"{datetime.now().month:02d}"
            
            erro_file = os.path.join(self.config.output_dir, "erros.txt")
            timestamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            
            with open(erro_file, "a", encoding="utf-8") as f:
                f.write(f"[{timestamp}] Competência: {mes_compet}/{ano_compet} - NSU: {nsu} - Chave: {chave} - Tipo: {tipo} - Erro: {descricao}\n")
            
            self.logger.error(f"Erro registrado: Competência {mes_compet}/{ano_compet} - NSU {nsu} - {tipo} - {descricao}")
            return True
        except Exception as e:
            self.logger.error(f"Falha ao registrar erro: {e}")
            return False
        
    def limpar_arquivo_erros(self):
        """Limpa o arquivo de erros no início de cada execução"""
        try:
            erro_file = os.path.join(self.config.output_dir, "erros.txt")
            if os.path.exists(erro_file):
                os.remove(erro_file)
                self.logger.info("Arquivo de erros anterior removido")
            with open(erro_file, "w", encoding="utf-8"):
                pass
            self.logger.info("Arquivo de erros limpo para nova execução")
        except Exception as e:
            self.logger.error(f"Erro ao limpar arquivo de erros: {e}")

    def carregar_nsu_competencia(self, nsu_competencia_file):
        """Carrega ou cria arquivo de competência no formato exato"""
        if os.path.exists(nsu_competencia_file):
            with open(nsu_competencia_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        else:
            # Cria estrutura exata conforme modelo
            return {"registros": {}}

    def obter_nsu_inicial_competencia(self, nsu_comp, ano, mes):
        """Obtém o NSU inicial para uma competência - SEMPRE do início"""
        registros = nsu_comp.get("registros", {})
        
        # Se já existe registro para esta competência, usar o nsu_inicial registrado
        if ano in registros and mes in registros[ano]:
            registro = registros[ano][mes]
            return registro.get("nsu_inicial", 1)
        
        # Buscar o maior NSU final de competências anteriores
        maior_nsu_anterior = 0
        competencia_atual = datetime(int(ano), int(mes), 1)  
        
        for a, meses_reg in registros.items():
            for m, dados in meses_reg.items():
                competencia_reg = datetime(int(a), int(m), 1) 
                if competencia_reg < competencia_atual:
                    nsu_final = dados.get("nsu_final", 0)
                    if nsu_final > maior_nsu_anterior:
                        maior_nsu_anterior = nsu_final
        
        # Se encontrou competências anteriores, começar do maior NSU + 1
        if maior_nsu_anterior > 0:
            return maior_nsu_anterior + 1
        
        # Primeira execução - começar do 1
        return 1

    def calcular_competencia_limite(self, ano_compet: str, mes_compet: str) -> tuple[str, str]:
        """Calcula a competência 6 meses após a escolhida (ano, mês)"""
        data_compet = datetime(int(ano_compet), int(mes_compet), 1)
        data_limite = data_compet + relativedelta(months=6)
        return str(data_limite.year), f"{data_limite.month:02d}"

    def deve_baixar_documento(self, ano_compet_doc, mes_compet_doc, ano_emissao_doc, mes_emissao_doc, 
                             ano_escolhido, mes_escolhido) -> tuple[bool, str]:
        """
        Verifica se o documento deve ser baixado baseado em competência OU emissão.
        
        Returns:
            tuple[bool, str]: (deve_baixar, motivo)
        """
        # Verificar por competência
        if ano_compet_doc == ano_escolhido and mes_compet_doc == mes_escolhido:
            return True, "COMPETÊNCIA"
        
        # Verificar por emissão
        if ano_emissao_doc == ano_escolhido and mes_emissao_doc == mes_escolhido:
            return True, "EMISSÃO"
        
        return False, "FORA DO PERÍODO"

    def verificar_documento_passado(self, ano_compet_doc, mes_compet_doc, ano_emissao_doc, mes_emissao_doc,
                                  ano_escolhido, mes_escolhido, nsu_item, chave) -> bool:
        """
        Verifica se o documento foi 'passado indevidamente' 
        (competência diferente da emissão mas uma delas é do mês escolhido).
        
        Returns:
            bool: True se é caso de auditoria (documento passado indevidamente)
        """
        # Criar datas para comparação
        try:
            data_compet = datetime(int(ano_compet_doc), int(mes_compet_doc), 1)
            data_emissao = datetime(int(ano_emissao_doc), int(mes_emissao_doc), 1)
            data_escolhida = datetime(int(ano_escolhido), int(mes_escolhido), 1)
            
            # Verificar se competência é diferente da emissão
            if data_compet != data_emissao:
                # Verificar se uma das datas é do mês escolhido
                competencia_do_mes = (data_compet.month == data_escolhida.month and 
                                     data_compet.year == data_escolhida.year)
                emissao_do_mes = (data_emissao.month == data_escolhida.month and 
                                 data_emissao.year == data_escolhida.year)
                
                # Se uma das datas é do mês escolhido, mas a outra não
                if competencia_do_mes != emissao_do_mes:
                    return True
        except Exception as e:
            self.logger.error(f"Erro ao verificar documento passado: {e}")
        
        return False

    ## ------------------------------------------------------------------------------
    ## Processo principal de download
    ## ------------------------------------------------------------------------------
    def run_competencia(self, ano_compet, mes_compet, nsu_competencia_file, write=None):
        """
        Executa download por competência específica (campo dCompet).
        Baixa documentos cuja competência OU emissão seja do mês escolhido.
        
        Args:
            ano_compet: Ano da competência (string)
            mes_compet: Mês da competência (string, formato "01" a "12")
            nsu_competencia_file: Caminho do arquivo de controle JSON
            write: Função callback para atualização de progresso
        
        Returns:
            int: Número de documentos baixados
        """
        if write is None:
            write = lambda msg, log=True: self.logger.info(msg) if log else None
        
        self.logger.info(f"Iniciando download VERIFICAÇÃO DUPLA para mês {mes_compet}/{ano_compet}")
        self.logger.info("Baixando documentos cuja COMPETÊNCIA OU EMISSÃO seja do mês escolhido")
        
        # Calcular competência limite (6 meses após)
        ano_limite, mes_limite = self.calcular_competencia_limite(ano_compet, mes_compet)
        self.logger.info(f"Buscando até NSU final da competência {mes_limite}/{ano_limite}")
        
        # Limpar e criar pastas da empresa
        self.limpar_pastas_empresa()
        self.criar_pastas_empresa()
        
        # Limpar arquivo de erros
        self.limpar_arquivo_erros()
        
        # Carregar/Criar arquivo de competência
        nsu_comp = self.carregar_nsu_competencia(nsu_competencia_file)
        
        # Obter NSU inicial baseado nos registros existentes
        nsu_inicial = self.obter_nsu_inicial_competencia(nsu_comp, ano_compet, mes_compet)
        nsu_atual = nsu_inicial - 1
        
        self.logger.info(f"NSU inicial: {nsu_inicial}")
        
        # Obter NSU limite (NSU final do 6º mês posterior)
        nsu_limite = None
        registros = nsu_comp.get("registros", {})
        
        if ano_limite in registros and mes_limite in registros[ano_limite]:
            nsu_limite = registros[ano_limite][mes_limite].get("nsu_final", None)
            if nsu_limite:
                self.logger.info(f"NSU limite definido: {nsu_limite} (final de {mes_limite}/{ano_limite})")
        
        # Contadores
        tent_erro = 0
        tent_post = 0  # Contador de competências posteriores ao limite
        documentos_baixados = 0
        auditorias_encontradas = 0
        
        # Dicionário para armazenar os intervalos por mês (emissão) durante esta execução
        intervalos_por_mes = {}  # chave: (ano, mes), valor: {"nsu_inicial": int, "nsu_final": int}

        # Configurar sessão
        with self.pfx_to_pem() as pem_cert:
            self.session = requests.Session()
            self.session.cert = pem_cert
            self.session.verify = True
            
            if self.config.download_pdf:
                pdf_dl = NFSePDFDownloader(self.session, self.config.timeout)
            
            try:
                while self.running() and tent_post < MAX_TENT:
                    
                    # Se já temos NSU limite e passamos dele, verificar se deve parar
                    if nsu_limite and nsu_atual > nsu_limite:
                        if tent_post >= MAX_TENT:
                            self.logger.info(f"Alcançado NSU final do 6º mês ({nsu_limite}). Encerrando busca.")
                            break
                    
                    url = f"{self.base_url}/{nsu_atual:020d}?cnpj={self.config.cnpj}"
                    
                    write(f"Consultando NSU: {nsu_atual}", log=False)
                    self.logger.info(f"Consultando NSU {nsu_atual}...")

                    try:
                        resp = self.session.get(url, timeout=self.config.timeout)
                    except requests.exceptions.RequestException as e:
                        status_code = getattr(e.response, 'status_code', 'N/A') if hasattr(e, 'response') else 'N/A'
                        error_msg = f"Erro de conexão no NSU {nsu_atual}: {e} (Status: {status_code})"
                        self.logger.error(error_msg)
                        self.registrar_erro(nsu_atual, "N/A", "CONEXÃO", f"{e} - Status: {status_code}", 
                                          ano_compet, mes_compet)
                        break
                    
                    if resp.status_code == 200:
                        resposta = resp.json()
                        if resposta.get("StatusProcessamento") == "DOCUMENTOS_LOCALIZADOS":
                            documentos = resposta.get("LoteDFe", [])
                            documentos = sorted(documentos, key=lambda d: int(d.get("NSU", 0)))
                            
                            for nfse in documentos:
                                if not self.running():
                                    break
                                    
                                nsu_item = int(nfse["NSU"])
                                chave = nfse["ChaveAcesso"]
                                arquivo_xml = nfse["ArquivoXml"]
                                
                                write(f"Processando NSU: {nsu_item}", log=False)
                                
                                try:
                                    # Processar XML
                                    xml_gzip = base64.b64decode(arquivo_xml)
                                    xml_bytes = gzip.decompress(xml_gzip)
                                    
                                    # Extrair COMPETÊNCIA e EMISSÃO
                                    ano_doc_compet, mes_doc_compet = self.extrair_competencia(xml_bytes)
                                    ano_doc_emissao, mes_doc_emissao = self.extrair_data_emissao(xml_bytes)
                                    
                                    self.logger.info(f"NSU {nsu_item} - Competência: {mes_doc_compet}/{ano_doc_compet}, "
                                                f"Emissão: {mes_doc_emissao}/{ano_doc_emissao}")
                                    
                                    # ATUALIZAR REGISTRO DE EMISSÃO no dicionário temporário
                                    chave_mes = (ano_doc_emissao, mes_doc_emissao)
                                    if chave_mes not in intervalos_por_mes:
                                        intervalos_por_mes[chave_mes] = {"nsu_inicial": nsu_item, "nsu_final": nsu_item}
                                    else:
                                        # Atualizar o menor nsu_inicial e maior nsu_final
                                        if nsu_item < intervalos_por_mes[chave_mes]["nsu_inicial"]:
                                            intervalos_por_mes[chave_mes]["nsu_inicial"] = nsu_item
                                        if nsu_item > intervalos_por_mes[chave_mes]["nsu_final"]:
                                            intervalos_por_mes[chave_mes]["nsu_final"] = nsu_item
                                    
                                    # VERIFICAÇÃO DUPLA: verificar se deve baixar (competência OU emissão)
                                    deve_baixar, motivo = self.deve_baixar_documento(
                                        ano_doc_compet, mes_doc_compet, 
                                        ano_doc_emissao, mes_doc_emissao,
                                        ano_compet, mes_compet
                                    )
                                    
                                    if deve_baixar:
                                        tent_erro = 0
                                        tent_post = 0
                                        
                                        # VERIFICAR SE É CASO DE AUDITORIA (documento "passado")
                                        if self.verificar_documento_passado(
                                            ano_doc_compet, mes_doc_compet,
                                            ano_doc_emissao, mes_doc_emissao,
                                            ano_compet, mes_compet,
                                            nsu_item, chave
                                        ):
                                            self.logger.warning(f"AUDITORIA: Documento NSU {nsu_item} 'passado indevidamente' - "
                                                              f"Competência: {mes_doc_compet}/{ano_doc_compet}, "
                                                              f"Emissão: {mes_doc_emissao}/{ano_doc_emissao}")
                                            auditorias_encontradas += 1
                                        
                                        # Determinar tipo do documento
                                        tipo_documento = self.determinar_tipo_documento(xml_bytes)
                                        self.logger.info(f"Documento {chave} classificado como: {tipo_documento} - Motivo: {motivo}")
                                        
                                        # Baixar arquivo
                                        pasta_tipo = os.path.join(self.config.output_dir, tipo_documento)
                                        filename = os.path.join(
                                            pasta_tipo, 
                                            f"{self.config.file_prefix}_NSU-{nsu_item}_{chave}.xml"
                                        )
                                        
                                        # Salvar XML
                                        with open(filename, "wb") as fxml:
                                            fxml.write(xml_bytes)
                                        
                                        documentos_baixados += 1
                                        write(f"XML baixado ({tipo_documento}): {chave} (NSU: {nsu_item}) - Motivo: {motivo}")
                                        
                                        # Baixar PDF se configurado
                                        if self.config.download_pdf:
                                            pdf_file = os.path.join(
                                                pasta_tipo,
                                                f"{self.config.file_prefix}_{nsu_item}_{chave}.pdf",
                                            )
                                            if pdf_dl.baixar(chave, pdf_file):
                                                self.logger.info(f"PDF baixado ({tipo_documento}): {chave}")
                                            else:
                                                self.logger.error(f"Falha ao baixar PDF: {chave}")
                                                self.registrar_erro(nsu_item, chave, "PDF", "Falha no download", 
                                                                  ano_compet, mes_compet)
                                    else:
                                        # Documento não é do mês escolhido (nem por competência, nem por emissão)
                                        self.logger.info(f"Documento fora do período: {mes_doc_compet}/{ano_doc_compet} - {mes_doc_emissao}/{ano_doc_emissao}")
                                        
                                        # Verificar se é posterior ao limite (pela EMISSÃO)
                                        doc_date_emissao = datetime(int(ano_doc_emissao), int(mes_doc_emissao), 1)
                                        doc_date_compet = datetime(int(ano_doc_compet), int(mes_doc_compet), 1)
                                        limite_date = datetime(int(ano_limite), int(mes_limite), 1)
                                        
                                        # Usar a data mais recente entre competência e emissão
                                        data_doc = max(doc_date_emissao, doc_date_compet)
                                        
                                        if data_doc > limite_date:
                                            # Competência/Emissão posterior ao limite
                                            tent_post += 1
                                            self.logger.info(f"Competência/Emissão posterior ao limite. Contador: {tent_post}/{MAX_TENT}")
                                        else:
                                            # Competência/Emissão anterior ou dentro do período
                                            tent_post = 0
                                    
                                except Exception as e:
                                    self.logger.error(f"Erro ao processar documento NSU {nsu_item}: {str(e)}")
                                    self.registrar_erro(nsu_item, chave, "XML", str(e), ano_compet, mes_compet)
                                    continue
                            
                            # Atualizar NSU atual
                            if documentos:
                                ultimo_nsu_processado = max(int(nfse["NSU"]) for nfse in documentos)
                                nsu_atual = ultimo_nsu_processado
                            else:
                                nsu_atual += 1
                            
                    elif resp.status_code in STATUS_STOP:
                        error_msg = f"Status de parada {resp.status_code} no NSU {nsu_atual}"
                        self.logger.info(error_msg)
                        self.registrar_erro(nsu_atual, "N/A", "HTTP", error_msg, ano_compet, mes_compet)
                        break

                    elif resp.status_code == 429:
                        error_msg = f"Rate limit (429) no NSU {nsu_atual}"
                        self.logger.info(error_msg)
                        self.registrar_erro(nsu_atual, "N/A", "MT.REQ (aumentar delay)", error_msg, 
                                          ano_compet, mes_compet)
                        break

                    else:
                        error_msg = f"Erro HTTP {resp.status_code} no NSU {nsu_atual}: {resp.text}"
                        self.logger.error(error_msg)
                        self.registrar_erro(nsu_atual, "N/A", "HTTP", error_msg, ano_compet, mes_compet)
                        tent_erro += 1
                        
                        nsu_atual += 1
                        
                        if tent_erro >= MAX_TENT:
                            self.logger.info(f"Máximo de tentativas de erro ({MAX_TENT}) atingido. Parando.")
                            break
                    
                    # Delay entre requisições
                    time.sleep(self.config.delay_seconds)
                    
            finally:
                # Atualizar o arquivo JSON com os intervalos coletados
                self.atualizar_arquivo_competencia(nsu_competencia_file, intervalos_por_mes, ano_compet, mes_compet)
                
                if self.session:
                    self.session.close()
                    self.session = None
            
            # Exibir resumo da execução
            self.logger.info("=" * 60)
            self.logger.info(f"RESUMO DA EXECUÇÃO - Mês {mes_compet}/{ano_compet}")
            self.logger.info(f"Total de documentos baixados: {documentos_baixados}")
            self.logger.info(f"Documentos 'passados indevidamente' detectados: {auditorias_encontradas}")
            self.logger.info(f"Primeiro NSU processado: {nsu_inicial}")
            self.logger.info(f"Último NSU processado: {nsu_atual}")
            
            # Exibir registros atualizados
            registros = nsu_comp.get("registros", {})
            if registros:
                self.logger.info("Registros atualizados:")
                for ano, meses in sorted(registros.items()):
                    for mes, dados in sorted(meses.items()):
                        self.logger.info(f"  {ano}-{mes}: NSU {dados.get('nsu_inicial', 0)} a {dados.get('nsu_final', 0)}")
            
            self.logger.info("=" * 60)
            
            return documentos_baixados

    ## ------------------------------------------------------------------------------
    ## Método para auditoria de competência
    ## ------------------------------------------------------------------------------
    def auditar_competencia(self, nsu_competencia_file, ano, mes):
        """
        Audita e corrige automaticamente a competência escolhida.
        Verifica consistência entre os registros de diferentes meses.
        """
        nsu_comp = self.carregar_nsu_competencia(nsu_competencia_file)
        registros = nsu_comp.get("registros", {})
        
        self.logger.info(f"=== AUDITORIA COMPETÊNCIA {mes}/{ano} ===")
        
        # Verificar se a competência existe
        if ano not in registros or mes not in registros[ano]:
            self.logger.warning(f"Competência {mes}/{ano} não encontrada para auditoria")
            return False
        
        correcoes_realizadas = False
        
        # 1. Coletar todos os registros em uma lista simples
        registros_lista = []
        
        for ano_reg, meses_reg in registros.items():
            for mes_reg, dados in meses_reg.items():
                registros_lista.append({
                    'ano': ano_reg,
                    'mes': mes_reg,
                    'nsu_inicial': dados.get("nsu_inicial", 0),
                    'nsu_final': dados.get("nsu_final", 0)
                })
        
        # 2. Encontrar o período mais recente (último)
        if registros_lista:
            # Converter para inteiros para encontrar o mais recente
            registros_com_data = []
            for reg in registros_lista:
                try:
                    ano_int = int(reg['ano'])
                    mes_int = int(reg['mes'])
                    registros_com_data.append({
                        'ano': reg['ano'],
                        'mes': reg['mes'],
                        'ano_int': ano_int,
                        'mes_int': mes_int,
                        'nsu_inicial': reg['nsu_inicial'],
                        'nsu_final': reg['nsu_final']
                    })
                except ValueError:
                    continue
            
            if registros_com_data:
                # Encontrar o mais recente (maior ano, maior mês)
                mais_recente = max(registros_com_data, key=lambda x: (x['ano_int'], x['mes_int']))
                self.logger.info(f"Período mais recente identificado: {mais_recente['ano']}-{mais_recente['mes']}")
                
                # Encontrar o penúltimo (segundo mais recente)
                registros_sem_mais_recente = [r for r in registros_com_data if r != mais_recente]
                if registros_sem_mais_recente:
                    penultimo = max(registros_sem_mais_recente, key=lambda x: (x['ano_int'], x['mes_int']))
                    
                    # VERIFICAÇÃO: Último período deve começar onde o anterior terminou +1
                    esperado_inicial = penultimo['nsu_final'] + 1
                    
                    if mais_recente['nsu_inicial'] != esperado_inicial:
                        self.logger.warning(
                            f"✗ ÚLTIMO PERÍODO {mais_recente['ano']}-{mais_recente['mes']}: "
                            f"nsu_inicial ({mais_recente['nsu_inicial']}) não começa após período anterior "
                            f"(deveria ser: {esperado_inicial} = {penultimo['ano']}-{penultimo['mes']}.nsu_final + 1)"
                        )
                        self.logger.info(f"✓ Corrigindo nsu_inicial do último período para {esperado_inicial}")
                        
                        # Atualizar no dicionário original
                        registros[mais_recente['ano']][mais_recente['mes']]["nsu_inicial"] = esperado_inicial
                        correcoes_realizadas = True
                    else:
                        self.logger.info(
                            f"✓ ÚLTIMO PERÍODO {mais_recente['ano']}-{mais_recente['mes']}: "
                            f"nsu_inicial ({mais_recente['nsu_inicial']}) está correto "
                            f"(continua de {penultimo['ano']}-{penultimo['mes']}.nsu_final + 1)"
                        )
        
        # 3. Verificar consistência interna (nsu_inicial <= nsu_final)
        for ano_reg, meses_reg in registros.items():
            for mes_reg, dados in meses_reg.items():
                nsu_inicial = dados.get("nsu_inicial", 0)
                nsu_final = dados.get("nsu_final", 0)
                
                if nsu_inicial > nsu_final:
                    self.logger.warning(
                        f"✗ INCONSISTÊNCIA INTERNA {ano_reg}-{mes_reg}: "
                        f"nsu_inicial ({nsu_inicial}) > nsu_final ({nsu_final})"
                    )
                    
                    # Corrigir trocando os valores
                    self.logger.info(f"✓ Trocando valores: nsu_inicial={nsu_final}, nsu_final={nsu_inicial}")
                    dados["nsu_inicial"] = nsu_final
                    dados["nsu_final"] = nsu_inicial
                    correcoes_realizadas = True
                else:
                    self.logger.info(
                        f"✓ {ano_reg}-{mes_reg}: "
                        f"nsu_inicial ({nsu_inicial}) <= nsu_final ({nsu_final})"
                    )
        
        # 4. Salvar correções se necessário
        if correcoes_realizadas:
            with open(nsu_competencia_file, 'w', encoding='utf-8') as f:
                json.dump(nsu_comp, f, indent=2, ensure_ascii=False)
            self.logger.info("✓ Arquivo de competência atualizado com correções")
            
            # Exibir registros corrigidos
            self.logger.info("Registros após correção:")
            for ano_reg, meses_reg in sorted(registros.items()):
                for mes_reg, dados in sorted(meses_reg.items()):
                    self.logger.info(
                        f"  {ano_reg}-{mes_reg}: "
                        f"NSU {dados.get('nsu_inicial', 0)} a {dados.get('nsu_final', 0)}"
                    )
        else:
            self.logger.info("✓ Nenhuma correção necessária - arquivo está consistente")
        
        self.logger.info("=====================================")
        
        return correcoes_realizadas

    def atualizar_arquivo_competencia(self, nsu_competencia_file, intervalos_por_mes, ano_processado=None, mes_processado=None):
        """
        Atualiza o arquivo JSON de competência com os intervalos coletados.
        
        Para o mês que está sendo processado (ano_processado/mes_processado), substitui
        completamente o registro existente. Para outros meses:
        1. Se o registro existente for inconsistente (nsu_inicial muito baixo), substitui
        2. Caso contrário, mescla os intervalos normalmente.
        """
        nsu_comp = self.carregar_nsu_competencia(nsu_competencia_file)
        
        # Primeiro, atualizar todos os registros
        for (ano, mes), intervalo in intervalos_por_mes.items():
            if ano not in nsu_comp["registros"]:
                nsu_comp["registros"][ano] = {}
            
            # Se este é o mês que está sendo processado, SUBSTITUIR o registro
            if ano == ano_processado and mes == mes_processado:
                nsu_comp["registros"][ano][mes] = {
                    "nsu_inicial": intervalo["nsu_inicial"],
                    "nsu_final": intervalo["nsu_final"]
                }
                self.logger.info(f"Registro do mês processado {mes}/{ano} substituído: NSU {intervalo['nsu_inicial']} a {intervalo['nsu_final']}")
            else:
                # Para outros meses, verificar consistência antes de mesclar
                if mes in nsu_comp["registros"][ano]:
                    registro_existente = nsu_comp["registros"][ano][mes]
                    
                    # Verificar se o registro existente é plausível
                    # Um registro é considerado implausível se:
                    # 1. O nsu_inicial existente for MUITO menor que o nsu_inicial coletado
                    # 2. E o nsu_final existente for MUITO maior que o nsu_final coletado
                    # (indica que o registro existente está completamente errado)
                    
                    existente_inicial = registro_existente.get("nsu_inicial", 0)
                    existente_final = registro_existente.get("nsu_final", 0)
                    coletado_inicial = intervalo["nsu_inicial"]
                    coletado_final = intervalo["nsu_final"]
                    
                    # Se o registro existente começar muito antes mas terminar muito depois,
                    # é provavelmente um registro antigo e incorreto
                    if (existente_inicial < coletado_inicial - 100 and  # Mais de 100 NSUs de diferença
                        existente_final > coletado_final + 100):        # em ambas as direções
                        self.logger.warning(f"Registro existente para {mes}/{ano} parece incorreto: "
                                        f"NSU {existente_inicial}-{existente_final}. "
                                        f"Substituindo por NSU {coletado_inicial}-{coletado_final}")
                        nsu_comp["registros"][ano][mes] = {"nsu_inicial": coletado_inicial, "nsu_final": coletado_final}
                    else:
                        # Mesclar normalmente - expandir o intervalo se necessário
                        novo_inicial = min(existente_inicial, coletado_inicial)
                        novo_final = max(existente_final, coletado_final)
                        registro_existente["nsu_inicial"] = novo_inicial
                        registro_existente["nsu_final"] = novo_final
                        self.logger.info(f"Registro do mês {mes}/{ano} expandido: NSU {novo_inicial} a {novo_final}")
                else:
                    # Criar novo registro
                    nsu_comp["registros"][ano][mes] = {"nsu_inicial": intervalo["nsu_inicial"], "nsu_final": intervalo["nsu_final"]}
                    self.logger.info(f"Novo registro para {mes}/{ano}: NSU {intervalo['nsu_inicial']} a {intervalo['nsu_final']}")
        
        # AGORA, após todas as atualizações, rodar auditoria de consistência
        # para garantir que todos os meses estejam sequenciais
        self.corrigir_consistencia_sequencial(nsu_comp)
        
        # Salvar arquivo atualizado
        with open(nsu_competencia_file, 'w', encoding='utf-8') as f:
            json.dump(nsu_comp, f, indent=2, ensure_ascii=False)
        
        self.logger.info("Arquivo de competência atualizado e consistência verificada.")

    def corrigir_consistencia_sequencial(self, nsu_comp):
        """
        Corrige a consistência sequencial dos registros, garantindo que cada mês
        comece onde o anterior terminou + 1.
        """
        self.logger.info("Verificando consistência sequencial dos registros...")
        
        # Coletar todos os registros em ordem temporal
        registros_ordenados = []
        
        for ano, meses in nsu_comp.get("registros", {}).items():
            for mes, dados in meses.items():
                try:
                    ano_int = int(ano)
                    mes_int = int(mes)
                    registros_ordenados.append({
                        'ano': ano,
                        'mes': mes,
                        'ano_int': ano_int,
                        'mes_int': mes_int,
                        'nsu_inicial': dados.get("nsu_inicial", 0),
                        'nsu_final': dados.get("nsu_final", 0)
                    })
                except ValueError:
                    continue
        
        # Ordenar por data
        registros_ordenados.sort(key=lambda x: (x['ano_int'], x['mes_int']))
        
        correcoes = 0
        
        # Corrigir sequência
        for i in range(1, len(registros_ordenados)):
            atual = registros_ordenados[i]
            anterior = registros_ordenados[i-1]
            
            # O NSU inicial do mês atual deve ser > NSU final do mês anterior
            esperado_inicial = anterior['nsu_final'] + 1
            
            if atual['nsu_inicial'] != esperado_inicial:
                # Verificar se a diferença é pequena (pode ser normal)
                if abs(atual['nsu_inicial'] - esperado_inicial) > 10:  # Mais de 10 NSUs de diferença
                    self.logger.warning(f"Inconsistência sequencial: {atual['ano']}-{atual['mes']} "
                                    f"começa em NSU {atual['nsu_inicial']}, mas deveria começar em {esperado_inicial} "
                                    f"(após {anterior['ano']}-{anterior['mes']} terminar em {anterior['nsu_final']})")
                    
                    # Corrigir apenas se o atual começar MUITO antes (indicando erro)
                    if atual['nsu_inicial'] < esperado_inicial - 10:
                        self.logger.info(f"Corrigindo nsu_inicial de {atual['ano']}-{atual['mes']} para {esperado_inicial}")
                        nsu_comp["registros"][atual['ano']][atual['mes']]["nsu_inicial"] = esperado_inicial
                        correcoes += 1
        
        if correcoes > 0:
            self.logger.info(f"Realizadas {correcoes} correções de consistência sequencial")
        else:
            self.logger.info("Todos os registros estão sequenciais")