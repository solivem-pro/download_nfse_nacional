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
class NFSeDownloader:
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
        """Obtém o NSU inicial para uma competência com lógica SIMPLIFICADA"""
        registros = nsu_comp.get("registros", {})
        
        # Se já existe registro para esta competência, continuar de onde parou
        if ano in registros and mes in registros[ano]:
            registro = registros[ano][mes]
            ultimo_verificado = registro.get("ultimo_verificado", 0)
            nsu_final = registro.get("nsu_final", 0)
            
            # Se último verificado é menor que nsu_final, continuar do próximo
            if ultimo_verificado < nsu_final:
                return ultimo_verificado + 1
            # Se já verificou tudo, recomeçar do início da competência
            else:
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
    def run_por_competencia(self, ano, mes, nsu_competencia_file, write=None):
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
        
        # Obter NSU inicial para a competência
        nsu_inicial = self.obter_nsu_inicial_competencia(nsu_comp, ano, mes)
        nsu_atual = nsu_inicial
        
        self.logger.info(f"NSU inicial para {mes}/{ano}: {nsu_inicial}")
        
        # Contador para controlar quando parar
        tent_erro = 0
        tent_post = 0
        
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
                                            f"{self.config.file_prefix}_{ano}-{mes}_{chave}.xml"
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
                                                f"{self.config.file_prefix}_{ano}-{mes}_{chave}.pdf",
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
                                        
                                        # Atualizar registro da competência diferente encontrada
                                        self.atualizar_competencia_diferente(
                                            nsu_competencia_file, nsu_comp, ano_doc, mes_doc, nsu_item
                                        )
                                    
                                except Exception as e:
                                    self.logger.error(f"Erro ao processar documento NSU {nsu_item}: {str(e)}")
                                    self.registrar_erro(nsu_item, chave, "XML", str(e))
                                    # Continuar processando outros documentos do lote
                                    continue
                            
                            if documentos:
                                ultimo_nsu_processado = max(int(nfse["NSU"]) for nfse in documentos)
                                nsu_atual = ultimo_nsu_processado + 1
                                self.logger.info(f"Lote processado. Próximo NSU: {nsu_atual}")
                            else:
                                nsu_atual += 1
                            
                            # Salvar progresso após processar lote
                            self.salvar_progresso_competencia(
                                nsu_competencia_file, nsu_comp, ano, mes, nsu_atual, 
                                documentos_baixados, nsu_inicial, primeiro_nsu_competencia
                            )
                            
                    elif resp.status_code in STATUS_STOP:
                        error_msg = f"Status de parada {resp.status_code} no NSU {nsu_atual}"
                        self.logger.info(error_msg)
                        self.registrar_erro(nsu_atual, "N/A", "HTTP", error_msg, ano, mes)
                        self.salvar_progresso_competencia(
                            nsu_competencia_file, nsu_comp, ano, mes, nsu_atual, 
                            documentos_baixados, nsu_inicial, primeiro_nsu_competencia
                        )
                        
                        # AUDITORIA APÓS STATUS DE PARADA
                        self.logger.info("Realizando auditoria após status de parada...")
                        self.auditar_competencia(nsu_competencia_file, ano, mes)
                        break

                    elif resp.status_code == 429:
                        error_msg = f"Status de parada {resp.status_code} no NSU {nsu_atual}"
                        self.logger.info(error_msg)
                        self.registrar_erro(nsu_atual, "N/A", "MT.REQ (aumentar delay)", error_msg, ano, mes)
                        self.salvar_progresso_competencia(
                            nsu_competencia_file, nsu_comp, ano, mes, nsu_atual, 
                            documentos_baixados, nsu_inicial, primeiro_nsu_competencia
                        )
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
                        self.salvar_progresso_competencia(
                            nsu_competencia_file, nsu_comp, ano, mes, nsu_atual, 
                            documentos_baixados, nsu_inicial, primeiro_nsu_competencia
                        )
                        
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
            
            # AUDITORIA FINAL - Sempre executar ao final do processo
            self.logger.info("Realizando auditoria final...")
            correcoes_finais = self.auditar_competencia(nsu_competencia_file, ano, mes)
            if correcoes_finais:
                self.logger.info("Foram realizadas correções finais no arquivo de competência")
            
            self.logger.info(f"Download concluído para competência {mes}/{ano}. Total de documentos: {documentos_baixados}")
            return documentos_baixados
        
    ## ------------------------------------------------------------------------------
    ## Auxiliares de registro e controle
    ## ------------------------------------------------------------------------------
    def log_estatisticas_competencias(self, competencias_encontradas, write):
        """Registra estatísticas das competências encontradas - APENAS LOG"""
        if competencias_encontradas:
            self.logger.info("=== ESTATÍSTICAS POR COMPETÊNCIA ===")
            for competencia, tipos in competencias_encontradas.items():
                self.logger.info(f"Competência {competencia}:")
                for tipo, quantidade in tipos.items():
                    self.logger.info(f"  {tipo}: {quantidade} documento(s)")
            self.logger.info("=====================================")
        else:
            self.logger.warning("Nenhum documento encontrado em nenhuma competência")

    def salvar_progresso_competencia(self, nsu_competencia_file, nsu_comp, ano, mes, nsu_atual, documentos_baixados, nsu_inicial=0, primeiro_nsu_encontrado=None):
        """Salva o progresso no arquivo de competência com lógica CORRIGIDA"""
        if "registros" not in nsu_comp:
            nsu_comp["registros"] = {}
        
        if ano not in nsu_comp["registros"]:
            nsu_comp["registros"][ano] = {}
        
        # Determinar NSU inicial real
        nsu_inicial_real = nsu_inicial
        if primeiro_nsu_encontrado is not None and primeiro_nsu_encontrado > 0:
            nsu_inicial_real = min(nsu_inicial_real, primeiro_nsu_encontrado) if nsu_inicial_real > 0 else primeiro_nsu_encontrado
        
        # E o último verificado também deve ser nsu_atual - 1
        nsu_final = nsu_atual - 1 if nsu_atual > nsu_inicial_real else nsu_inicial_real
        
        # Atualizar ou criar registro
        if mes not in nsu_comp["registros"][ano]:
            nsu_comp["registros"][ano][mes] = {
                "nsu_inicial": nsu_inicial_real,
                "nsu_final": nsu_final,
                "ultimo_verificado": nsu_final  
            }
        else:
            registro = nsu_comp["registros"][ano][mes]
            
            # Atualizar apenas se necessário
            registro["nsu_inicial"] = min(registro.get("nsu_inicial", nsu_inicial_real), nsu_inicial_real)
            registro["nsu_final"] = max(registro.get("nsu_final", 0), nsu_final)
            registro["ultimo_verificado"] = registro["nsu_final"]  
        
        # Log do estado
        self.logger.info(f"Progresso salvo: {mes}/{ano} - NSU inicial: {nsu_inicial_real}, NSU final: {nsu_final}")
        
        # Salvar arquivo
        with open(nsu_competencia_file, 'w', encoding='utf-8') as f:
            json.dump(nsu_comp, f, indent=2, ensure_ascii=False)

    def atualizar_competencia_diferente(self, nsu_competencia_file, nsu_comp, ano_doc, mes_doc, nsu_item):
        """Atualiza o registro para uma competência diferente encontrada"""
        if "registros" not in nsu_comp:
            nsu_comp["registros"] = {}
        
        if ano_doc not in nsu_comp["registros"]:
            nsu_comp["registros"][ano_doc] = {}
        
        # E o NSU final/último verificado também deve ser nsu_item
        nsu_inicial_correto = nsu_item
        nsu_final_correto = nsu_item
        
        # Se a competência diferente já existe no registro
        if mes_doc in nsu_comp["registros"][ano_doc]:
            registro = nsu_comp["registros"][ano_doc][mes_doc]
            # Atualizar nsu_inicial se encontrarmos um valor menor
            if nsu_inicial_correto < registro.get("nsu_inicial", 9999999):
                registro["nsu_inicial"] = nsu_inicial_correto
            # Atualizar NSU final se encontrarmos um valor maior
            if nsu_final_correto > registro.get("nsu_final", 0):
                registro["nsu_final"] = nsu_final_correto
            # Atualizar último verificado
            registro["ultimo_verificado"] = nsu_final_correto
        else:
            # Criar novo registro para a competência diferente
            nsu_comp["registros"][ano_doc][mes_doc] = {
                "nsu_inicial": nsu_inicial_correto,
                "nsu_final": nsu_final_correto,
                "ultimo_verificado": nsu_final_correto
            }
        
        # Salvar arquivo atualizado
        with open(nsu_competencia_file, 'w', encoding='utf-8') as f:
            json.dump(nsu_comp, f, indent=2, ensure_ascii=False)
        
        self.logger.info(f"Registro atualizado para competência {mes_doc}/{ano_doc} - NSU: {nsu_item}")

    def auditar_competencia(self, nsu_competencia_file, ano, mes):
        """Audita e corrige automaticamente a competência escolhida"""
        nsu_comp = self.carregar_nsu_competencia(nsu_competencia_file)
        registros = nsu_comp.get("registros", {})
        
        # Encontrar competência atual
        if ano not in registros or mes not in registros[ano]:
            self.logger.warning(f"Competência {mes}/{ano} não encontrada para auditoria")
            return False
            
        competencia_atual = registros[ano][mes]
        nsu_inicial_atual = competencia_atual.get("nsu_inicial", 0)
        nsu_final_atual = competencia_atual.get("nsu_final", 0)
        
        self.logger.info(f"=== AUDITORIA E CORREÇÃO COMPETÊNCIA {mes}/{ano} ===")
        self.logger.info(f"NSU inicial: {nsu_inicial_atual}, NSU final: {nsu_final_atual}")
        
        correcoes_realizadas = False
        
        # Encontrar competência anterior
        competencia_anterior = None
        data_atual = datetime(int(ano), int(mes), 1)  
        for a, meses in registros.items():
            for m, dados in meses.items():
                try:
                    data_reg = datetime(int(a), int(m), 1) 
                    if data_reg < data_atual:
                        if competencia_anterior is None or data_reg > competencia_anterior['data']:
                            competencia_anterior = {
                                'data': data_reg,
                                'ano': a,
                                'mes': m,
                                'nsu_final': dados.get("nsu_final", 0)
                            }
                except ValueError:
                    continue
        
        if competencia_anterior and competencia_anterior['nsu_final'] > 0:
            esperado_inicial = competencia_anterior['nsu_final'] + 1
            if nsu_inicial_atual != esperado_inicial:
                self.logger.warning(f"✗ NSU inicial INCORRETO - Esperado: {esperado_inicial} (Anterior {competencia_anterior['mes']}/{competencia_anterior['ano']}: {competencia_anterior['nsu_final']} + 1)")
                self.logger.info(f"✓ Corrigindo NSU inicial de {nsu_inicial_atual} para {esperado_inicial}")
                competencia_atual["nsu_inicial"] = esperado_inicial
                correcoes_realizadas = True
            else:
                self.logger.info(f"✓ NSU inicial CORRETO - Anterior {competencia_anterior['mes']}/{competencia_anterior['ano']}: {competencia_anterior['nsu_final']}")
        else:
            self.logger.info("✓ Primeira competência - Não há anterior para comparar")
        
        # Encontrar competência posterior
        competencia_posterior = None
        for a, meses in registros.items():
            for m, dados in meses.items():
                try:
                    data_reg = datetime(int(a), int(m), 1)  
                    if data_reg > data_atual:
                        if competencia_posterior is None or data_reg < competencia_posterior['data']:
                            competencia_posterior = {
                                'data': data_reg,
                                'ano': a,
                                'mes': m,
                                'nsu_inicial': dados.get("nsu_inicial", 0)
                            }
                except ValueError:
                    continue
        
        if competencia_posterior and competencia_posterior['nsu_inicial'] > 0:
            esperado_final = competencia_posterior['nsu_inicial'] - 1
            if nsu_final_atual != esperado_final:
                self.logger.warning(f"✗ NSU final INCORRETO - Esperado: {esperado_final} (Posterior {competencia_posterior['mes']}/{competencia_posterior['ano']}: {competencia_posterior['nsu_inicial']} - 1)")
                self.logger.info(f"✓ Corrigindo NSU final de {nsu_final_atual} para {esperado_final}")
                competencia_atual["nsu_final"] = esperado_final
                correcoes_realizadas = True
            else:
                self.logger.info(f"✓ NSU final CORRETO - Posterior {competencia_posterior['mes']}/{competencia_posterior['ano']}: {competencia_posterior['nsu_inicial']}")
        else:
            self.logger.info("✓ Última competência - Não há posterior para comparar")
        
        if competencia_atual.get("ultimo_verificado", 0) != competencia_atual.get("nsu_final", 0):
            self.logger.info(f"✓ Corrigindo último_verificado de {competencia_atual.get('ultimo_verificado', 0)} para {competencia_atual.get('nsu_final', 0)}")
            competencia_atual["ultimo_verificado"] = competencia_atual["nsu_final"]
            correcoes_realizadas = True
        
        # Salvar correções se necessário
        if correcoes_realizadas:
            with open(nsu_competencia_file, 'w', encoding='utf-8') as f:
                json.dump(nsu_comp, f, indent=2, ensure_ascii=False)
            self.logger.info("✓ Arquivo de competência atualizado com as correções")
        
        # Log dos valores finais
        nsu_inicial_corrigido = competencia_atual.get("nsu_inicial", 0)
        nsu_final_corrigido = competencia_atual.get("nsu_final", 0)
        self.logger.info(f"Valores finais - NSU inicial: {nsu_inicial_corrigido}, NSU final: {nsu_final_corrigido}")
        
        self.logger.info("=====================================")
        return correcoes_realizadas