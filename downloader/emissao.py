import os
import base64
import gzip
import logging
from datetime import datetime
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
## Classe de download, definições iniciais
## ------------------------------------------------------------------------------
class NFSeDownloaderEmissao:
    """Baixa os arquivos NFSe por competência"""

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
    def extrair_ano_mes(xml_bytes: bytes) -> tuple[str, str]:
        """Return the year and month from ``dhEmi`` or ``dhEvento``."""
        now = datetime.now()  
        try:
            root = ET.fromstring(xml_bytes)
            el = None
            for tag in ("dhEmi", "dhEvento", "DataEmissao"):
                el = root.find(f'.//{{*}}{tag}')
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
            
            # Namespace para evitar problemas com namespaces nos XMLs
            namespaces = {'ns': 'http://www.abrasf.org.br/nfse.xsd'}
            
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
        """Convert ``pfx_path`` to a temporary PEM file."""
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

    def registrar_erro(self, nsu: int, chave: str, tipo: str, descricao: str, ano: str = None, mes: str = None):
        """Registra erro no arquivo erros.txt da empresa - APPEND durante mesma execução"""
        try:
            # Se não foi fornecida competência específica, usar a competência atual do processamento
            if ano is None or mes is None:
                ano = str(datetime.now().year)
                mes = f"{datetime.now().month:02d}"
            
            erro_file = os.path.join(self.config.output_dir, "erros.txt")
            tent_errotamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            
            # MODO 'a' para ADICIONAR durante a mesma execução
            with open(erro_file, "a", encoding="utf-8") as f:
                f.write(f"[{tent_errotamp}] Competência: {mes}/{ano} - NSU: {nsu} - Chave: {chave} - Tipo: {tipo} - Erro: {descricao}\n")
            
            self.logger.error(f"Erro registrado: Competência {mes}/{ano} - NSU {nsu} - {tipo} - {descricao}")
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
        """Carrega ou cria arquivo de competência"""
        if os.path.exists(nsu_competencia_file):
            with open(nsu_competencia_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        else:
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

    ## ------------------------------------------------------------------------------
    ## Processos de download da competência
    ## ------------------------------------------------------------------------------
    def atualizar_arquivo_competencia(self, nsu_competencia_file, intervalos_por_mes):
        """Atualiza o arquivo JSON de competência com os intervalos coletados"""
        nsu_comp = self.carregar_nsu_competencia(nsu_competencia_file)
        
        for (ano, mes), intervalo in intervalos_por_mes.items():
            if ano not in nsu_comp["registros"]:
                nsu_comp["registros"][ano] = {}
            
            if mes not in nsu_comp["registros"][ano]:
                nsu_comp["registros"][ano][mes] = {
                    "nsu_inicial": intervalo["nsu_inicial"],
                    "nsu_final": intervalo["nsu_final"]
                }
            else:
                # Mesclar intervalos: pegar o menor nsu_inicial e o maior nsu_final
                registro = nsu_comp["registros"][ano][mes]
                registro["nsu_inicial"] = min(registro["nsu_inicial"], intervalo["nsu_inicial"])
                registro["nsu_final"] = max(registro["nsu_final"], intervalo["nsu_final"])
        
        # Salvar arquivo atualizado
        with open(nsu_competencia_file, 'w', encoding='utf-8') as f:
            json.dump(nsu_comp, f, indent=2, ensure_ascii=False)
        
        self.logger.info("Arquivo de competência atualizado com os intervalos coletados.")

    def run_emissao(self, ano, mes, nsu_competencia_file, write=None):
        """Executa download por competência específica - APENAS da competência escolhida"""
        if write is None:
            write = lambda msg, log=True: self.logger.info(msg) if log else None
        
        self.logger.info(f"Iniciando download para competência {mes}/{ano}")

        # Limpar e criar pastas da empresa
        self.limpar_pastas_empresa()
        self.criar_pastas_empresa()
        
        # Limpar arquivo de erros no início de cada execução
        self.limpar_arquivo_erros()
        
        # Carregar/Criar arquivo de competência
        nsu_comp = self.carregar_nsu_competencia(nsu_competencia_file)
        
        # AUDITORIA INICIAL - Corrigir possíveis inconsistências antes de começar
        self.logger.info("Realizando auditoria inicial...")
        correcoes_iniciais = self.auditar_competencia(nsu_competencia_file, ano, mes)
        if correcoes_iniciais:
            self.logger.info("Foram realizadas correções iniciais no arquivo de competência")
            # Recarregar o arquivo após correções
            nsu_comp = self.carregar_nsu_competencia(nsu_competencia_file)
        
        # Obter NSU inicial para a competência - SEMPRE do início
        nsu_inicial = self.obter_nsu_inicial_competencia(nsu_comp, ano, mes)
        nsu_atual = nsu_inicial - 1 
        
        self.logger.info(f"NSU inicial para {mes}/{ano}: {nsu_inicial} (sempre do início)")
        
        # Contador para controlar quando parar
        tent_erro = 0
        tent_post = 0
        
        # Dicionário para armazenar os intervalos por mês durante esta execução
        intervalos_por_mes = {}  # chave: (ano, mes), valor: {"nsu_inicial": int, "nsu_final": int}
        
        # Configurar sessão
        with self.pfx_to_pem() as pem_cert:
            self.session = requests.Session()
            self.session.cert = pem_cert
            self.session.verify = True
            
            if self.config.download_pdf:
                pdf_dl = NFSePDFDownloader(self.session, self.config.timeout)
            
            documentos_baixados = 0
            primeiro_nsu_competencia = None
            
            try:
                while self.running() and tent_post < MAX_TENT:
                    
                    if tent_post >= MAX_TENT:
                        self.logger.info(f"Encontradas {tent_post} notas de competência posterior seguidas. Parando busca.")
                        break
                    
                    url = f"{self.base_url}/{nsu_atual:020d}?cnpj={self.config.cnpj}"
                    
                    write(f"Consultando NSU: {nsu_atual}", log=False)
                    self.logger.info(f"Consultando a partir do NSU {nsu_atual}...")

                    try:
                        resp = self.session.get(url, timeout=self.config.timeout)
                    except requests.exceptions.RequestException as e:
                        status_code = getattr(e.response, 'status_code', 'N/A') if hasattr(e, 'response') else 'N/A'
                        error_msg = f"Erro de conexão no NSU {nsu_atual}: {e} (Status: {status_code})"
                        self.logger.error(error_msg)
                        self.registrar_erro(nsu_atual, "N/A", "CONEXÃO", f"{e} - Status: {status_code}", ano, mes)
                        
                        # AUDITORIA EM CASO DE ERRO - Verificar se precisa corrigir algo
                        self.logger.info("Realizando auditoria após erro de conexão...")
                        self.auditar_competencia(nsu_competencia_file, ano, mes)
                        break
                    
                    if resp.status_code == 200:
                        resposta = resp.json()
                        if resposta.get("StatusProcessamento") == "DOCUMENTOS_LOCALIZADOS":
                            documentos = resposta.get("LoteDFe", [])
                            documentos = sorted(documentos, key=lambda d: int(d.get("NSU", 0)))
                            
                            # Flag para verificar se encontrou algum documento da competência neste lote
                            encontrou_documento_competencia = False
                            
                            for nfse in documentos:
                                if not self.running():
                                    break
                                    
                                nsu_item = int(nfse["NSU"])
                                chave = nfse["ChaveAcesso"]
                                arquivo_xml = nfse["ArquivoXml"]
                                
                                write(f"Processando NSU: {nsu_item}", log=False)
                                self.logger.info(f"Processando NSU {nsu_item}...")
                                
                                try:
                                    # Processar XML
                                    xml_gzip = base64.b64decode(arquivo_xml)
                                    xml_bytes = gzip.decompress(xml_gzip)
                                    
                                    # Verificar competência do documento
                                    ano_doc, mes_doc = self.extrair_ano_mes(xml_bytes)
                                    self.logger.info(f"Documento NSU {nsu_item} - Competência extraída: {mes_doc}/{ano_doc}")
                                    
                                    # ATUALIZAR INTERVALO PARA ESTE MÊS no dicionário temporário
                                    chave_mes = (ano_doc, mes_doc)
                                    if chave_mes not in intervalos_por_mes:
                                        intervalos_por_mes[chave_mes] = {
                                            "nsu_inicial": nsu_item,
                                            "nsu_final": nsu_item
                                        }
                                    else:
                                        # Atualizar o menor nsu_inicial e maior nsu_final
                                        if nsu_item < intervalos_por_mes[chave_mes]["nsu_inicial"]:
                                            intervalos_por_mes[chave_mes]["nsu_inicial"] = nsu_item
                                        if nsu_item > intervalos_por_mes[chave_mes]["nsu_final"]:
                                            intervalos_por_mes[chave_mes]["nsu_final"] = nsu_item
                                    
                                    # VERIFICAR SE É DA COMPETÊNCIA ESCOLHIDA
                                    if ano_doc == ano and mes_doc == mes:
                                        # Competência correta - baixar
                                        encontrou_documento_competencia = True
                                        tent_erro = 0  # Resetar contador  de erros 
                                        tent_post = 0  # Resetar contador de notas posteriores
                                        
                                        # Determinar tipo do documento
                                        tipo_documento = self.determinar_tipo_documento(xml_bytes)
                                        self.logger.info(f"Documento {chave} classificado como: {tipo_documento}")
                                        
                                        # Registrar primeiro NSU da competência
                                        if primeiro_nsu_competencia is None:
                                            primeiro_nsu_competencia = nsu_item
                                            self.logger.info(f"Primeiro documento da competência {mes}/{ano} encontrado no NSU: {primeiro_nsu_competencia}")
                                        
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
                                        write(f"XML baixado ({tipo_documento}): {chave} (NSU: {nsu_item})")
                                        
                                        # Baixar PDF se configurado
                                        if self.config.download_pdf:
                                            pdf_file = os.path.join(
                                                pasta_tipo,
                                                f"{self.config.file_prefix}_{nsu_atual}_{chave}.pdf",
                                            )
                                            if pdf_dl.baixar(chave, pdf_file):
                                                self.logger.info(f"PDF baixado ({tipo_documento}): {chave}")
                                            else:
                                                self.logger.error(f"Falha ao baixar PDF: {chave}")
                                                self.registrar_erro(nsu_item, chave, "PDF", "Falha no download", ano_doc, mes_doc)
                                    
                                    else:
                                        # Competência diferente - apenas atualizar registro
                                        self.logger.info(f"Documento de competência diferente: {mes_doc}/{ano_doc} (NSU: {nsu_item}) - Atualizando registro apenas")
                                        
                                        doc_date = datetime(int(ano_doc), int(mes_doc), 1)  
                                        target_date = datetime(int(ano), int(mes), 1)  
                                        
                                        if doc_date > target_date:
                                            # É uma competência posterior
                                            tent_post += 1
                                            self.logger.info(f"Competência posterior encontrada. Contador: {tent_post}/{MAX_TENT}")
                                        else:
                                            # É uma competência anterior - IGNORAR E CONTINUAR BUSCA
                                            # Não incrementar tent_post, apenas continuar procurando
                                            tent_post = 0
                                            self.logger.info(f"Competência anterior encontrada ({mes_doc}/{ano_doc}). Continuando busca...")
                                    
                                except Exception as e:
                                    self.logger.error(f"Erro ao processar documento NSU {nsu_item}: {str(e)}")
                                    self.registrar_erro(nsu_item, chave, "XML", str(e))
                                    # Continuar processando outros documentos do lote
                                    continue
                            
                            if documentos:
                                ultimo_nsu_processado = max(int(nfse["NSU"]) for nfse in documentos)
                                nsu_atual = ultimo_nsu_processado
                                self.logger.info(f"Lote processado. Próximo NSU: {nsu_atual}")
                            else:
                                nsu_atual += 1
                            
                    elif resp.status_code in STATUS_STOP:
                        error_msg = f"Status de parada {resp.status_code} no NSU {nsu_atual}"
                        self.logger.info(error_msg)
                        self.registrar_erro(nsu_atual, "N/A", "HTTP", error_msg, ano, mes)
                        
                        # AUDITORIA APÓS STATUS DE PARADA
                        self.logger.info("Realizando auditoria após status de parada...")
                        self.auditar_competencia(nsu_competencia_file, ano, mes)
                        break

                    elif resp.status_code == 429:
                        error_msg = f"Status de parada {resp.status_code} no NSU {nsu_atual}"
                        self.logger.info(error_msg)
                        self.registrar_erro(nsu_atual, "N/A", "MT.REQ (aumentar delay)", error_msg, ano, mes)
                        break

                    else:
                        error_msg = f"Erro HTTP {resp.status_code} no NSU {nsu_atual}: {resp.text}"
                        self.logger.error(error_msg)
                        self.registrar_erro(nsu_atual, "N/A", "HTTP", error_msg)
                        # Incrementar contador quando há erro HTTP
                        tent_erro += 1
                        self.logger.info(f"Erro HTTP. Tentativas consecutivas: {tent_erro}/{MAX_TENT}")
                        
                        # Avançar NSU mesmo com erro
                        nsu_atual += 1
                        
                        # Parar apenas se exceder o máximo de tentativas de erro
                        if tent_erro >= MAX_TENT:
                            self.logger.info(f"Máximo de tentativas de erro ({MAX_TENT}) atingido. Parando.")
                            break
                    
                    # Delay entre requisições
                    time.sleep(self.config.delay_seconds)
                    
            finally:
                if self.session:
                    self.session.close()
                    self.session = None
            
            # Atualizar o arquivo JSON com todos os intervalos coletados
            self.atualizar_arquivo_competencia(nsu_competencia_file, intervalos_por_mes)
            
            # AUDITORIA FINAL - Sempre executar ao final do processo
            self.logger.info("Realizando auditoria final...")
            correcoes_finais = self.auditar_competencia(nsu_competencia_file, ano, mes)
            if correcoes_finais:
                self.logger.info("Foram realizadas correções finais no arquivo de competência")
            
            # Exibir resumo dos intervalos coletados
            if intervalos_por_mes:
                self.logger.info("Intervalos coletados por mês:")
                for (ano_mes, mes_mes), intervalo in sorted(intervalos_por_mes.items()):
                    self.logger.info(f"  {mes_mes}/{ano_mes}: NSU {intervalo['nsu_inicial']} a {intervalo['nsu_final']}")
            
            self.logger.info(f"Download concluído para competência {mes}/{ano}. Total de documentos: {documentos_baixados}")
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