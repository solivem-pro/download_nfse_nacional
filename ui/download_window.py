import json
import os
import shutil
import logging
import tkinter as tk
import threading
import time  # Mover esta importação para cá
import traceback
import zipfile
import win32com.client as win32
from pathlib import Path
import pythoncom
from tkinter import ttk, messagebox, filedialog
from datetime import datetime, timedelta
## Módulos auxiliares
from config.config import DIRETORIOS, ROOT_DIR, Config
from config.utils import formatar_cnpj
from downloader.emissao import NFSeDownloaderEmissao
from downloader.competencia import NFSeDownloaderCompetencia
from ui.ui_basic import PopupProcessamento, notificar_windows, modal_window, scrolled_treeview, buttons_frame, back_window
from config.config import Config
from config.json_handler import carregar_cadastros
logger = logging.getLogger(__name__)

## ------------------------------------------------------------------------------
## Interface principal de Download
## ------------------------------------------------------------------------------
class DownloadUI:
    """Janela de seleção e download de NFSe"""
    
    def __init__(self, parent):
        self.parent = parent
        self.win = None
        self.data = None
        self.chaves_cadastros = []
        self.tree = None
        self.processo_ativo = False
        self.resultados = []
        self.contador_nfse_global = 0
        self.empresas_selecionadas = []
        self.indice_cnpj = {}  # Adicionar este
        
        self._setup_ui()

    def _setup_ui(self):
        """Configura a interface principal"""
        self.win = modal_window(self.parent.root, "Download - Download NFSe Nacional", 500, 500)

        # Carregar cadastros existentes
        self.data = carregar_cadastros()
        if not self.data:
            messagebox.showerror("Erro", "Nenhum cadastro encontrado!")
            back_window(self.win, self.parent.root)
            return

        # Criar índice rápido de CNPJ para exportação
        self._criar_indice_cnpj()
        
        self._criar_treeview()
        self._criar_filtros()
        self._criar_botoes()
        self._atualizar_lista()

        # Focar na janela e no Treeview para permitir atalhos de teclado
        self.win.focus_set()
        self.tree.focus_set()

    def _criar_treeview(self):
        """Cria a treeview de empresas"""
        columns_config = [
            ("cod", "Código", 80, "center", "int"),
            ("empresa", "Empresa", 200, "center", "string"), 
            ("cnpj", "CNPJ", 150, "center", "string") 
        ]
        
        self.tree, scrollbar, frame_tree = scrolled_treeview(
            self.win, 
            columns_config, 
            height=15,
            show="headings"
        )

        # Configurar tags para formatação condicional
        self.tree.tag_configure('vencido', foreground='red', font=('TkDefaultFont', 9, 'bold'))
        self.tree.tag_configure('vencido_selecionado', foreground='white', background='red', font=('TkDefaultFont', 9, 'bold'))
        self.tree.tag_configure('normal', foreground='black')
        self.tree.tag_configure('selecionado', foreground='white', background='blue')

        # Configurar eventos
        self.tree.bind('<Button-1>', self._on_treeview_click)
        self.tree.bind('<<TreeviewSelect>>', self._on_tree_select)
        self.tree.bind('<KeyPress>', self._on_key_press)

    def _criar_filtros(self):
        """Cria os filtros de ano e mês"""
        frame_filtros = tk.Frame(self.win)
        frame_filtros.pack(fill=tk.X, padx=10, pady=10)

        # Ano
        tk.Label(frame_filtros, text="Ano:").grid(row=0, column=0, padx=5, sticky="w")
        anos = [str(ano) for ano in range(datetime.now().year - 5, datetime.now().year + 1)]
        self.combo_ano = ttk.Combobox(frame_filtros, values=anos, state="readonly", width=10)
        self.combo_ano.set(str(datetime.now().year))
        self.combo_ano.grid(row=0, column=1, padx=5)

        # Mês
        tk.Label(frame_filtros, text="Mês:").grid(row=0, column=2, padx=5, sticky="w")
        meses = [
            ("01", "Janeiro"), ("02", "Fevereiro"), ("03", "Março"),
            ("04", "Abril"), ("05", "Maio"), ("06", "Junho"),
            ("07", "Julho"), ("08", "Agosto"), ("09", "Setembro"),
            ("10", "Outubro"), ("11", "Novembro"), ("12", "Dezembro")
        ]
        hoje = datetime.now()
        mes_anterior = hoje.replace(day=1) - timedelta(days=1)
        self.combo_mes = ttk.Combobox(frame_filtros, values=[m[0] for m in meses], state="readonly", width=10)
        self.combo_mes.set(str(mes_anterior.month).zfill(2))
        self.combo_mes.grid(row=0, column=3, padx=5)

    def _criar_botoes(self):
        """Cria os botões de ação"""
        botoes_config = [
            {"text": "Selec. Todos", "width": 10, "command": self._selecionar_todos},
            {"text": "Baixar", "width": 10, "state": tk.DISABLED, "command": self._baixar_nfse},
            {"text": "Exportar", "width": 10, "state": tk.DISABLED, "command": self._exportar_nfse},
        ]
        
        frame_buttons, botoes = buttons_frame(self.win, botoes_config)
        self.btn_baixar = botoes["Baixar"]
        self.btn_exportar = botoes["Exportar"]

        # Botão voltar à direita
        btn_close = tk.Button(frame_buttons, text="Voltar", width=10, command= lambda: back_window(self.win, self.parent.root))
        btn_close.pack(side=tk.RIGHT, padx=20)

    def _verificar_certificado_vencido(self, cadastro):
        """Verifica se o certificado está vencido"""
        vencimento_str = cadastro.get('venc', '')
        if not vencimento_str:
            return False
            
        try:
            data_venc = datetime.strptime(vencimento_str, "%d/%m/%Y")
            data_hoje = datetime.now()
            
            # Considera vencido se a data atual é maior que a data de vencimento
            return data_hoje > data_venc
        except ValueError:
            return False

    def _atualizar_lista(self):
        """Atualiza a lista de empresas"""
        self.tree.delete(*self.tree.get_children())
        self.chaves_cadastros = []
        
        # Log do carregamento
        logger.info(f"Carregando cadastros do arquivo: {DIRETORIOS['cadastros_json']}")
        logger.info(f"Estrutura do JSON: {list(self.data.keys())}")
        
        # Buscar todas as chaves que começam com "cadastro_"
        chaves_cadastro = [k for k in self.data.keys() if k.startswith('cadastro_') and k != 'cadastro_0']
        logger.info(f"Chaves de cadastro encontradas: {chaves_cadastro}")
        
        # Ordenar cadastros por código
        cadastros_ordenados = []
        for key in chaves_cadastro:
            cadastro = self.data[key]
            if 'cod' in cadastro:
                cadastros_ordenados.append((key, cadastro))
        
        # Ordenar por código
        cadastros_ordenados.sort(key=lambda item: int(item[1]["cod"]))
        
        logger.info(f"Total de cadastros válidos encontrados: {len(cadastros_ordenados)}")
        
        for key, cadastro in cadastros_ordenados:
            # Verificar se o certificado está vencido
            certificado_vencido = self._verificar_certificado_vencido(cadastro)
            
            # Determinar a tag baseada no status do certificado
            tag = 'vencido' if certificado_vencido else 'normal'
            
            self.tree.insert("", tk.END, values=(
                cadastro['cod'],
                cadastro['empresa'],
                formatar_cnpj(cadastro['cnpj'])
            ), tags=(tag,))
            
            self.chaves_cadastros.append(key)
            logger.info(f"Adicionado à lista: Código {cadastro['cod']} - {cadastro['empresa']} - Vencido: {certificado_vencido}")
        
        logger.info(f"Cadastros carregados com sucesso. Total: {len(cadastros_ordenados)}")

    def _on_treeview_click(self, event):
        """Alterna a seleção individual dos itens ao clicar"""
        # Verificar se o clique foi na região do cabeçalho
        region = self.tree.identify_region(event.x, event.y)
        
        if region == "heading":
            # Se foi no cabeçalho, não faz nada - deixa a ordenação padrão funcionar
            return
        
        item = self.tree.identify_row(event.y)
        if item:
            # Verificar se o item tem certificado vencido
            tags = self.tree.item(item, 'tags')
            certificado_vencido = 'vencido' in tags
            
            if certificado_vencido:
                # Se está vencido, aplicar tag de selecionado vencido
                if item in self.tree.selection():
                    self.tree.selection_remove(item)
                    self.tree.item(item, tags=('vencido',))
                else:
                    self.tree.selection_add(item)
                    self.tree.item(item, tags=('vencido_selecionado',))
            else:
                # Comportamento normal para certificados válidos
                if item in self.tree.selection():
                    self.tree.selection_remove(item)
                else:
                    self.tree.selection_add(item)
                    
        self._atualizar_estado_botoes()
        return "break"

    def _atualizar_estado_botoes(self):
        """Atualiza o estado dos botões baseado na seleção e status dos certificados"""
        selecionados = self.tree.selection()
        
        if not selecionados:
            self.btn_baixar.config(state=tk.DISABLED)
            self.btn_exportar.config(state=tk.DISABLED)
            return

        # Verificar se há pelo menos uma empresa com certificado válido selecionada
        empresas_validas_selecionadas = False
        empresas_vencidas_selecionadas = []
        
        for item in selecionados:
            tags = self.tree.item(item, 'tags')
            if 'vencido' in tags or 'vencido_selecionado' in tags:
                values = self.tree.item(item, 'values')
                empresas_vencidas_selecionadas.append(values[1])  # Nome da empresa
            else:
                empresas_validas_selecionadas = True

        # Habilitar/desabilitar botões
        if empresas_validas_selecionadas:
            self.btn_baixar.config(state=tk.NORMAL)
        else:
            self.btn_baixar.config(state=tk.DISABLED)
            
        # Exportar sempre está disponível (não depende do certificado)
        self.btn_exportar.config(state=tk.NORMAL)

    def _on_tree_select(self, event):
        """Atualiza tags quando a seleção muda"""
        # Atualizar tags para refletir seleção
        for item in self.tree.get_children():
            tags = self.tree.item(item, 'tags')
            if item in self.tree.selection() and 'vencido' in tags:
                self.tree.item(item, tags=('vencido_selecionado',))
            elif 'vencido_selecionado' in tags and item not in self.tree.selection():
                self.tree.item(item, tags=('vencido',))
                
        self._atualizar_estado_botoes()

    def _on_key_press(self, event):
        """Manipula atalhos de teclado"""
        if event.state & 0x4 and event.keysym.lower() == 'a':  # Ctrl+A
            self._selecionar_todos()
            return "break"

    def _selecionar_todos(self):
        """Seleciona todos os itens no Treeview, mas mantém formatação de certificados vencidos"""
        for item in self.tree.get_children():
            tags = self.tree.item(item, 'tags')
            if 'vencido' in tags:
                self.tree.selection_add(item)
                self.tree.item(item, tags=('vencido_selecionado',))
            else:
                self.tree.selection_add(item)
                
        self._atualizar_estado_botoes()

    def _buscar_cadastro_empresa(self, cod_empresa):
        """Busca cadastro da empresa com logging detalhado"""
        logger.info(f"Buscando cadastro para empresa código: {cod_empresa}")
        
        try:
            # Log da estrutura completa do JSON para debug
            logger.info(f"Estrutura completa do JSON carregado: {list(self.data.keys())}")
            logger.info(f"Total de cadastros reportado: {self.data.get('cadastros', 'N/A')}")
            
            # Buscar em todas as chaves que começam com "cadastro_"
            chaves_cadastro = [k for k in self.data.keys() if k.startswith('cadastro_')]
            logger.info(f"Chaves de cadastro encontradas: {chaves_cadastro}")
            
            for key in chaves_cadastro:
                cadastro = self.data[key]
                logger.info(f"Verificando chave {key}: código {cadastro.get('cod', 'N/A')}")
                
                # Comparar como string para evitar problemas de tipo
                if str(cadastro.get('cod', '')) == str(cod_empresa):
                    empresa_encontrada = cadastro.get('empresa', 'N/A')
                    cnpj_encontrado = cadastro.get('cnpj', 'N/A')
                    vencimento = cadastro.get('venc', 'N/A')
                    logger.info(f"✅ Cadastro encontrado: {empresa_encontrada} (CNPJ: {cnpj_encontrado}, Venc: {vencimento})")
                    logger.info(f"Detalhes completos: {cadastro}")
                    return cadastro
            
            # Se não encontrou, log mais detalhado
            logger.error(f"❌ Cadastro não encontrado para código: {cod_empresa}")
            logger.error(f"Códigos disponíveis: {[self.data[k].get('cod', 'N/A') for k in chaves_cadastro]}")
            logger.error(f"Tipo do código buscado: {type(cod_empresa)}, valor: {cod_empresa}")
            return None
            
        except Exception as e:
            logger.error(f"Erro inesperado ao buscar cadastro para {cod_empresa}: {str(e)}")
            logger.error(f"Traceback: {traceback.format_exc()}")
            return None

## ------------------------------------------------------------------------------
## Exportar pacotes
## ------------------------------------------------------------------------------
    def _exportar_nfse(self):
        """Função para exportar arquivos .zip das empresas selecionadas com base no save_mode configurado"""
        selecionados = self.tree.selection()
        if not selecionados:
            return
            
        # Carregar configurações
        config = Config.load(DIRETORIOS['config_json'])
        save_mode = getattr(config, 'save_mode')
        
        logger.info(f"Iniciando exportação com save_mode: {save_mode}")
        
        destino = filedialog.askdirectory(title="Selecione a pasta para exportar os arquivos ZIP")
        if not destino:
            return
        
        empresas_exportadas = []
        empresas_nao_encontradas = []
        empresas_erro_renomeacao = []
        
        # Criar pasta temporária para cópias renomeadas
        pasta_temp = os.path.join(DIRETORIOS['temp'], "export_temp")
        if os.path.exists(pasta_temp):
            shutil.rmtree(pasta_temp)
        os.makedirs(pasta_temp, exist_ok=True)
        
        # Processar cada empresa selecionada
        for item in selecionados:
            values = self.tree.item(item, 'values')
            cod_empresa = values[0]
            nome_empresa = values[1]
            
            nome_arquivo_origem = f"{cod_empresa}.zip"
            caminho_origem = os.path.join(DIRETORIOS['notas'], nome_arquivo_origem)
            
            if not os.path.exists(caminho_origem):
                empresas_nao_encontradas.append(f"[{cod_empresa}] {nome_empresa}")
                continue
            
            try:
                # Determinar nome final baseado no save_mode
                if save_mode == 'CNPJ':
                    # Usar índice para busca rápida
                    cnpj_limpo = self.indice_cnpj.get(cod_empresa)
                    if cnpj_limpo:
                        nome_arquivo_destino = f"{cnpj_limpo}.zip"
                    else:
                        # Fallback: buscar no cadastro
                        cadastro = self._buscar_cadastro_empresa(cod_empresa)
                        if cadastro and 'cnpj' in cadastro:
                            cnpj_formatado = cadastro['cnpj'].replace('.', '').replace('/', '').replace('-', '')
                            nome_arquivo_destino = f"{cnpj_formatado}.zip"
                            # Atualizar índice
                            self.indice_cnpj[cod_empresa] = cnpj_formatado
                        else:
                            logger.warning(f"CNPJ não encontrado para empresa {cod_empresa}. Usando código como fallback.")
                            nome_arquivo_destino = nome_arquivo_origem
                else:
                    # Modo código (default)
                    nome_arquivo_destino = nome_arquivo_origem
                
                # Criar cópia temporária renomeada
                caminho_temp = os.path.join(pasta_temp, nome_arquivo_destino)
                shutil.copy2(caminho_origem, caminho_temp)
                
                # Copiar da pasta temporária para o destino final
                caminho_destino = os.path.join(destino, nome_arquivo_destino)
                shutil.copy2(caminho_temp, caminho_destino)
                
                # Adicionar à lista de exportadas com formato adequado
                if save_mode == 'cnpj':
                    # Mostrar transformação apenas se CNPJ foi encontrado
                    cnpj_usado = self.indice_cnpj.get(cod_empresa, '')
                    if cnpj_usado:
                        empresas_exportadas.append(f"[{cod_empresa}] {nome_empresa} -> {cnpj_usado}.zip")
                    else:
                        empresas_exportadas.append(f"[{cod_empresa}] {nome_empresa}")
                else:
                    empresas_exportadas.append(f"[{cod_empresa}] {nome_empresa}")
                    
                logger.info(f"Exportado: {cod_empresa} -> {nome_arquivo_destino}")
                
            except Exception as e:
                logger.error(f"Erro ao exportar empresa {cod_empresa}: {e}")
                empresas_erro_renomeacao.append(f"[{cod_empresa}] {nome_empresa}")
        
        # Limpar pasta temporária
        try:
            shutil.rmtree(pasta_temp)
            logger.info(f"Pasta temporária limpa: {pasta_temp}")
        except Exception as e:
            logger.warning(f"Erro ao limpar pasta temporária: {e}")
        
        # Construir mensagem de resultado
        mensagem = f"Exportação de arquivos ZIP ({save_mode}):\n\n"
        
        if empresas_exportadas:
            mensagem += f"✅ {len(empresas_exportadas)} arquivo(s) exportado(s):\n"
            mensagem += "\n".join(f"  • {emp}" for emp in empresas_exportadas[:10])  # Limitar a 10 itens
            if len(empresas_exportadas) > 10:
                mensagem += f"\n  ... e mais {len(empresas_exportadas) - 10} arquivos"
        
        if empresas_nao_encontradas:
            if empresas_exportadas:
                mensagem += "\n\n"
            mensagem += f"❌ {len(empresas_nao_encontradas)} arquivo(s) não encontrado(s):\n"
            mensagem += "\n".join(f"  • {emp}" for emp in empresas_nao_encontradas[:5])
            if len(empresas_nao_encontradas) > 5:
                mensagem += f"\n  ... e mais {len(empresas_nao_encontradas) - 5} arquivos"
        
        if empresas_erro_renomeacao:
            if empresas_exportadas or empresas_nao_encontradas:
                mensagem += "\n\n"
            mensagem += f"⚠️ {len(empresas_erro_renomeacao)} erro(s) de processamento:\n"
            mensagem += "\n".join(f"  • {emp}" for emp in empresas_erro_renomeacao)
        
        if not empresas_exportadas and not empresas_nao_encontradas and not empresas_erro_renomeacao:
            mensagem += "⚠️ Nenhuma empresa processada."
        
        # Adicionar nota sobre CNPJ formatado
        if save_mode == 'cnpj':
            mensagem += "\n\n📝 Nota: CNPJs exportados sem pontuação (formato: 12345678000195)"
        
        messagebox.showinfo("Exportação Concluída", mensagem)
        logger.info(f"Exportação concluída. Modo: {save_mode}, Exportados: {len(empresas_exportadas)}")

    def _criar_indice_cnpj(self):
        """Cria índice rápido código->CNPJ para exportação"""
        self.indice_cnpj = {}
        
        try:
            for key in self.data.keys():
                if key.startswith('cadastro_'):
                    cadastro = self.data[key]
                    cod = str(cadastro.get('cod', ''))
                    cnpj = cadastro.get('cnpj', '')
                    if cod and cnpj:
                        # Remover pontuação do CNPJ
                        cnpj_limpo = cnpj.replace('.', '').replace('/', '').replace('-', '')
                        self.indice_cnpj[cod] = cnpj_limpo
        except Exception as e:
            logger.error(f"Erro ao criar índice CNPJ: {e}")

    ## ------------------------------------------------------------------------------
    ## Processos do download e compactação
    ## ------------------------------------------------------------------------------
    def _processar_apos_download(self, cod_empresa, cnpj_empresa):
        """Processa as etapas pós-download"""
        try:
            pasta_empresa = os.path.join(DIRETORIOS['notas'], str(cod_empresa))
            xlsm_files = [f for f in os.listdir(pasta_empresa) if f.endswith('.xlsm')]
            
            if not xlsm_files:
                logger.warning(f"Nenhum arquivo .xlsm encontrado para empresa {cod_empresa}")
                return False
                
            xlsm_path = os.path.join(pasta_empresa, xlsm_files[0])
            
            macro_executada = self._executar_macro_vba(xlsm_path, 'ImportarTodosXMLs')
            
            if not macro_executada:
                logger.warning(f"Falha ao executar macro para empresa {cod_empresa}")
            
            zip_path = os.path.join(DIRETORIOS['notas'], f"{cod_empresa}.zip")
            compactacao_sucesso = self._compactar_pasta_empresa(pasta_empresa, zip_path)
            
            if not compactacao_sucesso:
                logger.error(f"Falha ao compactar pasta da empresa {cod_empresa}")
                return False
                
            return True
            
        except Exception as e:
            logger.error(f"Erro no processamento pós-download: {e}")
            return False

    def _executar_macro_vba(self, caminho_xlsm, nome_macro):
        """Executa macro do Excel em modo oculto"""
        excel = None
        workbook = None
        try:
            pythoncom.CoInitialize()
            
            excel = win32.DispatchEx("Excel.Application")
            excel.Visible = 0
            excel.DisplayAlerts = 0
            excel.AskToUpdateLinks = 0
            excel.ScreenUpdating = 0
            excel.EnableEvents = 0
            excel.Interactive = 0
            
            time.sleep(0.5)  # Usar time.sleep em vez de datetime.time.sleep
            
            workbook = excel.Workbooks.Open(caminho_xlsm, ReadOnly=False)
            excel.Run(nome_macro)
            workbook.Save()
            workbook.Close()
            excel.Quit()
            
            logging.info(f"Macro {nome_macro} executada com sucesso")
            return True
            
        except Exception as e:
            logging.error(f"Erro ao executar macro {nome_macro}: {e}")
            try:
                if workbook:
                    workbook.Close(SaveChanges=False)
                if excel:
                    excel.Quit()
            except:
                pass
            return False
        finally:
            try:
                if workbook:
                    workbook = None
                if excel:
                    excel = None
            except:
                pass
            try:
                pythoncom.CoUninitialize()
            except:
                pass

    def _compactar_pasta_empresa(self, pasta_origem, caminho_zip):
        """Compacta toda a pasta incluindo subpastas"""
        try:
            pasta_origem = Path(pasta_origem)
            
            with zipfile.ZipFile(caminho_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for root, dirs, files in os.walk(pasta_origem):
                    for dir_name in dirs:
                        dir_path = Path(root) / dir_name
                        rel_path = dir_path.relative_to(pasta_origem)
                        zip_info = zipfile.ZipInfo(str(rel_path).replace("\\", "/") + "/")
                        zipf.writestr(zip_info, "")
                    
                    for file_name in files:
                        file_path = Path(root) / file_name
                        rel_path = file_path.relative_to(pasta_origem)
                        zipf.write(file_path, rel_path)
                        
            logger.info(f"Pasta {pasta_origem} compactada para {caminho_zip}")
            return True
            
        except Exception as e:
            logger.error(f"Erro ao compactar pasta {pasta_origem}: {e}")
            return False

    def _atualizar_contador_nfse(self, incremento=1):
        """Atualiza o contador global de NFSe"""
        self.contador_nfse_global += incremento
        if hasattr(self, 'popup') and self.popup and hasattr(self.popup, 'winfo_exists') and self.popup.winfo_exists():
            self.win.after(0, lambda: self.popup.atualizar_contador_nfse(self.contador_nfse_global))

    def _baixar_nfse(self):
        """Função principal para baixar NFSe das empresas selecionadas"""
        selecionados = self.tree.selection()
        if not selecionados:
            return
            
        ano = self.combo_ano.get()
        mes = self.combo_mes.get()
        
        # Log adicional para debug
        logger.info(f"Iniciando download para {len(selecionados)} empresas selecionadas")
        logger.info(f"Competência: {mes}/{ano}")

        config_global = Config.load(DIRETORIOS['config_json'])
        
        self.empresas_selecionadas = []
        empresas_ignoradas_vencimento = []
        
        for item in selecionados:
            values = self.tree.item(item, 'values')
            cod_empresa = values[0]
            nome_empresa = values[1]
            
            logger.info(f"Processando seleção: código {cod_empresa}, nome {nome_empresa}")
            
            # Buscar cadastro da empresa
            cadastro = self._buscar_cadastro_empresa(cod_empresa)
            
            if cadastro is None:
                logger.error(f"Falha ao encontrar cadastro para empresa {cod_empresa} - {nome_empresa}")
                messagebox.showerror("Erro", f"Cadastro não encontrado para a empresa {cod_empresa}")
                continue
            
            # Verificar se o certificado está vencido
            if self._verificar_certificado_vencido(cadastro):
                logger.warning(f"Certificado vencido para empresa {nome_empresa}. Ignorando no download.")
                # Armazenar tanto o código quanto o nome
                empresas_ignoradas_vencimento.append({
                    'cod': cod_empresa,
                    'nome': nome_empresa
                })
                continue
                
            self.empresas_selecionadas.append({
                'cod': cod_empresa,
                'nome': nome_empresa,
                'cadastro': cadastro
            })
        
        # Avisar sobre empresas ignoradas por certificado vencido
        if empresas_ignoradas_vencimento:
            mensagem_aviso = f"Total de {len(empresas_ignoradas_vencimento)} empresas com certificados vencidos ignoradas:"
            for empresa_info in empresas_ignoradas_vencimento:
                mensagem_aviso += f"\n • [{empresa_info['cod']}] {empresa_info['nome']}"
            messagebox.showwarning("Certificados Vencidos", mensagem_aviso)
            logger.warning(mensagem_aviso)
        
        if not self.empresas_selecionadas:
            logger.error("Nenhuma empresa válida para processar")
            messagebox.showwarning("Nenhuma Empresa Válida", 
                                "Todas as empresas selecionadas possuem certificados vencidos. "
                                "Atualize os certificados antes de fazer o download.")
            return

        self.popup = PopupProcessamento(
            self.win, 
            titulo="Baixando - Download NFS-e Nacional", 
            texto=f"Baixando NFSe para {len(self.empresas_selecionadas)} empresa(s) válida(s)..."
        )
        
        self.resultados = []
        self.processo_ativo = True
        self.contador_nfse_global = 0
        
        # Iniciar thread de download
        thread_download = threading.Thread(target=self._processo_download, daemon=True)
        thread_download.start()

    def _processo_download(self):
        """Processo de download em thread separada"""
        try:
            total_empresas = len(self.empresas_selecionadas)
            self.total_empresas = total_empresas  # Armazena o total
            logger.info(f"Iniciando download para {total_empresas} empresas válidas")
        
            for i, empresa in enumerate(self.empresas_selecionadas, 1):
                self.empresa_atual_index = i  # Armazena o índice atual
                logger.info(f"Processando empresa {i}/{total_empresas}: {empresa['nome']}")
            
                if self.processo_ativo:
                    try:
                        # Atualiza o popup com o índice da empresa e NSU zero
                        self.win.after(0, lambda idx=i, total=total_empresas: self.popup.atualizar_contador(idx, total, 0))
                    except Exception as e:
                        logger.warning(f"Erro ao atualizar popup: {e}")
                
                resultado = self._baixar_empresa(empresa, self.combo_ano.get(), self.combo_mes.get())
                
                logger.info(f"Resultado para [{empresa['cod']}] {empresa['nome']}: {resultado['documentos']} documentos, {resultado['erros']} erros")
                
                self.resultados.append(resultado)

                if self.processo_ativo and resultado['erros'] == 0:
                    try:                       
                        processamento_ok = self._processar_apos_download(empresa['cod'], empresa['cadastro']['cnpj'])
                        if processamento_ok:
                            logger.info(f"Processamento pós-download concluído para [{empresa['cod']}] {empresa['nome']}")
                        else:
                            logger.warning(f"Problemas no processamento pós-download para [{empresa['cod']}] {empresa['nome']}")
                    except Exception as e:
                        logger.error(f"Erro no processamento pós-download para [{empresa['cod']}] {empresa['nome']}: {e}")
                elif not self.processo_ativo:
                    logger.warning(f"Processamento pós-download pulado para [{empresa['cod']}] {empresa['nome']} - processo cancelado")
                else:
                    logger.warning(f"Pulando processamento pós-download para [{empresa['cod']}] {empresa['nome']} devido a erros no download")
                    
        except Exception as e:
            logger.error(f"Erro no processo de download: {e}")
        finally:
            logger.info("Processo de download finalizado")
            for resultado in self.resultados:
                logger.info(f"Resultado final - [{resultado.get('cod', 'N/A')}] {resultado['empresa']}: {resultado['documentos']} documentos")
            
            if self.processo_ativo:
                self.win.after(0, self._finalizar_processo)

    def _baixar_empresa(self, empresa, ano, mes):
        """Baixa NFSe para uma empresa específica"""
        documentos_baixados = 0
        erros = 0
        
        try:
            cod_empresa = empresa['cod']
            nome_empresa = empresa['nome']
            cadastro = empresa['cadastro']
            
            # Carregar configurações
            config = Config.load(DIRETORIOS['config_json'])
            consult_mode = getattr(config, 'consult_mode')  # Default para Emissão
            
            # Verificação de segurança - garantir que o certificado não está vencido
            if self._verificar_certificado_vencido(cadastro):
                error_msg = f"Certificado vencido para {nome_empresa}. Download cancelado."
                logger.error(error_msg)
                return {
                    'cod': cod_empresa,
                    'empresa': nome_empresa,
                    'documentos': 0,
                    'erros': 1,
                    'mensagem': error_msg
                }
            
            cert_path = os.path.join(ROOT_DIR, cadastro['cert_path'])
            cert_pass = cadastro['cert_pass']
            cnpj = cadastro['cnpj']
                        
            pasta_empresa = os.path.join(DIRETORIOS['notas'], cod_empresa)
            if not os.path.exists(pasta_empresa):
                logger.error(f"Pasta da empresa {cod_empresa} não encontrada")
                return {
                    'cod': cod_empresa,
                    'empresa': nome_empresa,
                    'documentos': 0,
                    'erros': 1,
                    'mensagem': f"Pasta da empresa {cod_empresa} não encontrada. Refazer cadastro."
                }            
            
            arquivo_controle = os.path.join(pasta_empresa, 'nsu_competencia.json')
            
            config_empresa = Config.load(DIRETORIOS['config_json'])
            config_empresa.cert_path = cert_path
            config_empresa.cert_pass = cert_pass
            config_empresa.cnpj = cnpj
            config_empresa.output_dir = pasta_empresa
            
            # ALTERAÇÃO: Instanciar o downloader correto conforme o modo
            if consult_mode == 'Emissão':
                from downloader.emissao import NFSeDownloaderEmissao
                downloader = NFSeDownloaderEmissao(config_empresa)
            else:  # Competência
                from downloader.competencia import NFSeDownloaderCompetencia
                downloader = NFSeDownloaderCompetencia(config_empresa)
            
            # Variável local para armazenar o NSU atual
            nsu_atual_local = 0
            
            def write_progress(msg, log=True):
                nonlocal nsu_atual_local
                
                if log:
                    timestamp = datetime.now().strftime("%H:%M:%S")
                    mensagem_completa = f"[{timestamp}] {cod_empresa}: {msg}"
                    logger.info(mensagem_completa)
                
                if "XML baixado" in msg:
                    self._atualizar_contador_nfse(1)
                
                # Extrai o NSU da mensagem se disponível
                if "NSU:" in msg or "Consultando NSU:" in msg:
                    try:
                        # Exemplo de mensagem: "Consultando NSU: 12345"
                        if "NSU:" in msg:
                            nsu_str = msg.split("NSU:")[1].strip().split()[0]
                            nsu_atual_local = int(nsu_str)
                    except (IndexError, ValueError):
                        pass  # Ignora se não conseguir extrair o NSU
                
                # Atualiza o contador com o NSU atual
                if hasattr(self, 'popup') and self.popup and hasattr(self.popup, 'winfo_exists') and self.popup.winfo_exists():
                    self.win.after(0, lambda: self.popup.atualizar_contador(
                        self.empresa_atual_index,
                        self.total_empresas,
                        nsu_atual_local
                    ))
            
            logger.info(f"Iniciando download para {nome_empresa} - Competência: {mes}/{ano}")
            logger.info(f"Modo de consulta: {consult_mode}")
            
            try:
                if consult_mode == 'Emissão':
                    documentos_baixados = downloader.run_emissao(
                        ano=ano,
                        mes=mes,
                        nsu_competencia_file=arquivo_controle,
                        write=write_progress
                    )
                elif consult_mode == 'Competência':
                    documentos_baixados = downloader.run_competencia(
                        ano_compet=ano,
                        mes_compet=mes,
                        nsu_competencia_file=arquivo_controle,
                        write=write_progress
                    )
                else:
                    error_msg = f"Modo de consulta desconhecido: {consult_mode}"
                    logger.error(error_msg)
                    return {
                        'cod': cod_empresa,
                        'empresa': nome_empresa,
                        'documentos': 0,
                        'erros': 1,
                        'mensagem': error_msg
                    }
                
                logger.info(f"Download concluído para [{cod_empresa}] {nome_empresa}: {documentos_baixados} documentos")
                write_progress(f"Download concluído. Total de documentos baixados: {documentos_baixados}")
                
                return {
                    'cod': cod_empresa,
                    'empresa': nome_empresa,
                    'documentos': documentos_baixados,
                    'erros': 0,
                    'mensagem': f"Sucesso: {documentos_baixados} documentos baixados"
                }
                    
            except Exception as e:
                error_msg = f"Erro durante download para {nome_empresa}: {str(e)}"
                logger.error(error_msg)
                logger.exception("Detalhes do erro:")  # Log completo do stack trace
                return {
                    'cod': cod_empresa,
                    'empresa': nome_empresa,
                    'documentos': 0,
                    'erros': 1,
                    'mensagem': error_msg
                }
                
        except Exception as e:
            error_msg = f"Erro ao baixar para {empresa['nome']}: {str(e)}"
            logger.error(error_msg)
            logger.exception("Detalhes do erro:")
            return {
                'cod': empresa.get('cod', 'N/A'),
                'empresa': empresa.get('nome', 'N/A'),
                'documentos': 0,
                'erros': 1,
                'mensagem': error_msg
            }

    def _finalizar_processo(self):
        """Finaliza o processo de download"""
        self.processo_ativo = False
        
        try:
            self.popup.finalizar()
        except Exception as e:
            logger.warning(f"Erro ao fechar popup: {e}")
            
        message="Processo de download finalizado, pronto para exportar arquivos"
        notificar_windows(message)
        self._exibir_resumo_download()      

    def _exibir_resumo_download(self):
        """Exibe popup com resumo do download"""
        if not self.resultados:
            return
        
        total_empresas = len(self.resultados)
        total_documentos = sum(r['documentos'] for r in self.resultados)
        total_erros = sum(r['erros'] for r in self.resultados)
        
        mensagem = f"Download concluído para {total_empresas} empresa(s)\n\n"
        mensagem += f"Total de documentos baixados: {total_documentos}\n"
        mensagem += f"Total de erros: {total_erros}\n\n"
        mensagem += "Detalhes por empresa:\n"
        
        for resultado in self.resultados:
            status = "✅" if resultado['erros'] == 0 else "❌"
            # Adicionar o código entre colchetes antes do nome da empresa
            mensagem += f"\n{status} [{resultado.get('cod', 'N/A')}] {resultado['empresa']} - Notas: {resultado['documentos']} - Erros: {resultado['erros']}"
        
        messagebox.showinfo("Resumo do Download", mensagem)

