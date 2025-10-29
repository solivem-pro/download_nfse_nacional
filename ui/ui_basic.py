import logging
import tkinter as tk
from tkinter import ttk
from pathlib import Path
from typing import Optional
from config.config import DIRETORIOS
logger = logging.getLogger(__name__)

## ------------------------------------------------------------------------------
## Config popups e janela
## ------------------------------------------------------------------------------
def centralizar(janela, largura: int = 400, altura: int = 300) -> None:
    """Centraliza uma janela na tela."""
    janela.update_idletasks()
    screen_width = janela.winfo_screenwidth()
    screen_height = janela.winfo_screenheight()
    x = (screen_width - largura) // 2
    y = (screen_height - altura) // 2
    janela.geometry(f"{largura}x{altura}+{x}+{y}")
    janela.resizable(False, False)

def modal_window(parent, titulo, largura, altura):
    """Cria uma janela modal padronizada"""
    win = tk.Toplevel(parent)
    win.title(titulo)
        # Tenta definir o mesmo ícone da janela principal
    try:
        from config.config import DIRETORIOS
        if DIRETORIOS['icone'].exists():
            if str(DIRETORIOS['icone']).endswith('.ico'):
                win.iconbitmap(str(DIRETORIOS['icone']))
    except Exception:
        pass  # Ignora erros no ícone de janelas secundárias
    centralizar(win, largura, altura)
    
    win.transient(parent)
    win.grab_set()
    win.focus_set()
    
    def on_closing():
        win.grab_release()
        win.destroy()
    
    win.protocol("WM_DELETE_WINDOW", on_closing)
    return win

def back_window(win, root):
    """Fecha a janela modal e retorna o foco para a janela principal"""
    win.destroy()
    root.grab_release()
    root.lift()
    root.focus_set()

def _set_window_icon(self):
    """Define o ícone da janela principal"""
    try:
        # Método 1: iconbitmap para arquivos .ico
        if DIRETORIOS['icone'].exists() and str(DIRETORIOS['icone']).endswith('.ico'):
            self.root.iconbitmap(str(DIRETORIOS['icone']))
            logging.info(f"Ícone definido: {DIRETORIOS['icone']}")
            
        else:
            logging.warning(f"Arquivo de ícone não encontrado: {DIRETORIOS['icone']}")
            
    except Exception as e:
        logging.error(f"Erro ao carregar ícone: {str(e)}")

## ----------------------------------------------------------------------------------------------
## Treeview genérico com scrollbar
## ----------------------------------------------------------------------------------------------
def scrolled_treeview(parent, columns_config, height=15, show='headings'):
    """
    Cria um treeview com scrollbar configurado de forma genérica.
    
    Args:
        parent: Widget pai
        columns_config: Lista de tuplas (id, texto, largura, anchor)
        height: Altura do treeview
        show: Configuração de exibição
    
    Returns:
        tuple: (treeview, scrollbar)
    """
    frame_tree = tk.Frame(parent)
    frame_tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
    columns = [col[0] for col in columns_config]
    tree = ttk.Treeview(frame_tree, columns=columns, show=show, height=height)
    
    # Configurar cabeçalhos e colunas
    for col_id, heading, width, anchor in columns_config:
        tree.heading(col_id, text=heading, anchor=anchor)
        tree.column(col_id, width=width, anchor=anchor)
    
    # Scrollbar
    scrollbar = ttk.Scrollbar(frame_tree, orient=tk.VERTICAL, command=tree.yview)
    tree.configure(yscrollcommand=scrollbar.set)
    
    tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    return tree, scrollbar, frame_tree

def refresh_treeview(tree, dados, chaves, formatadores=None):
    """
    Atualiza um treeview com dados de forma genérica.
    
    Args:
        tree: Treeview a ser atualizado
        dados: Lista de dicionários ou tuplas com os dados
        chaves: Lista de chaves para referenciar os dados
        formatadores: Dicionário com funções de formatação para cada coluna
    """
    tree.delete(*tree.get_children())
    
    if formatadores is None:
        formatadores = {}
    
    for item_data in dados:
        if isinstance(item_data, dict):
            values = []
            for col in tree['columns']:
                valor = item_data.get(col, '')
                # Aplica formatador se existir
                if col in formatadores:
                    valor = formatadores[col](valor)
                values.append(valor)
        else:
            values = item_data
            
        tree.insert('', tk.END, values=values)
    
    return chaves

def formatar_dados_treeview(dados, formatadores):
    """
    Aplica formatadores a dados do treeview.
    
    Args:
        dados: Dados originais
        formatadores: Dicionário de funções de formatação
    
    Returns:
        list: Dados formatados
    """
    if isinstance(dados, dict):
        return {k: formatadores.get(k, lambda x: x)(v) for k, v in dados.items()}
    elif isinstance(dados, (list, tuple)):
        return [formatadores.get(i, lambda x: x)(v) for i, v in enumerate(dados)]
    return dados

## ----------------------------------------------------------------------------------------------
## UI helpers
## ----------------------------------------------------------------------------------------------
class ToolTip:
    """Cria uma tooltip ao passar o mouse sobre um widget."""
    
    def __init__(self, widget: tk.Widget, text: str):
        self.widget = widget
        self.text = text
        self.tip_window = None
        widget.bind("<Enter>", self.show_tip)
        widget.bind("<Leave>", self.hide_tip)

    def show_tip(self, event: Optional[tk.Event] = None) -> None:
        if self.tip_window or not self.text:
            return
        
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + 20
        
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        
        label = tk.Label(
            tw, text=self.text, justify="left",
            background="#ffffe0", relief="solid", borderwidth=1,
            font=("Segoe UI", 9)
        )
        label.pack(ipadx=5, ipady=3)

    def hide_tip(self, event: Optional[tk.Event] = None) -> None:
        if self.tip_window:
            self.tip_window.destroy()
            self.tip_window = None

def notificar_windows(mensagem: str) -> None:
    """Versão simplificada e robusta para notificações do Windows."""
    try:
        from plyer import notification
        
        # Verifica se o arquivo de ícone existe
        icone_path = Path(DIRETORIOS['icone'])
        app_icon = str(icone_path) if icone_path.exists() else None
        
        notification.notify(
            title="NFS-e Nacional",
            message=mensagem,
            app_name="Download NFS-e Nacional",
            app_icon=app_icon,  # Usa None se o ícone não existir
            timeout=5  # Aumentado para melhor visibilidade
        )
    except Exception as e:
        # Fallback silencioso - não quebra a aplicação se a notificação falhar
        print(f"Notificação falhou: {e}")

## Popup de processamento
class PopupProcessamento(tk.Toplevel):
    """Popup modal para exibir progresso de operações."""

    def __init__(self, parent, titulo: str = "Processando...", texto: str = "Aguarde o processamento"):
        super().__init__(parent)
        self.parent = parent
        self.title(titulo)
        centralizar(self, 400, 180)
        self._set_window_icon() 
        # Torna a janela modal
        self.transient(parent)
        self.grab_set()
        self.resizable(False, False)
        
        self._setup_ui(texto)
        self.cancelado = False
        
        # Impede fechamento pelo X
        self.protocol("WM_DELETE_WINDOW", self._prevent_close)

    def _set_window_icon(self):
        """Define o ícone da janela de popup - MÉTODO ADICIONADO"""
        try:
            if DIRETORIOS['icone'].exists() and str(DIRETORIOS['icone']).endswith('.ico'):
                self.iconbitmap(str(DIRETORIOS['icone']))
                logging.info(f"Ícone do popup definido: {DIRETORIOS['icone']}")
            else:
                logging.warning(f"Arquivo de ícone não encontrado: {DIRETORIOS['icone']}")
        except Exception as e:
            logging.error(f"Erro ao carregar ícone do popup: {str(e)}")

    def _setup_ui(self, texto: str) -> None:
        """Configura a interface do usuário do popup."""
        frame_principal = tk.Frame(self, padx=20, pady=20)
        frame_principal.pack(fill=tk.BOTH, expand=True)
        
        # Label de texto
        tk.Label(
            frame_principal, 
            text=texto,
            font=("Arial", 10)
        ).pack(pady=(0, 15))
        
        # Label do contador de empresas
        self.label_contador = tk.Label(
            frame_principal,
            text="Processando: 0/0 - NSU: 0",
            font=("Arial", 9),
            fg="blue"
        )
        self.label_contador.pack(pady=(0, 5))
        
        # Label do contador de NFSe
        self.label_nfse = tk.Label(
            frame_principal,
            text="NFSe baixadas: 0",
            font=("Arial", 9),
            fg="green"
        )
        self.label_nfse.pack(pady=(0, 10))
        
        # Barra de progresso
        self.progressbar = ttk.Progressbar(
            frame_principal,
            mode='indeterminate',
            length=350
        )
        self.progressbar.pack(pady=(0, 15))
        self.progressbar.start(10)

    def _prevent_close(self) -> None:
        """Impede o fechamento do popup pelo usuário."""
        pass

    def atualizar_contador(self, atual: int, total: int, nsu: int) -> None:
        """Atualiza o contador de empresas processadas."""
        self._safe_label_update(
            self.label_contador, 
            f"Processando: {atual}/{total} - NSU: {nsu}"
        )

    def atualizar_contador_nfse(self, nfse_baixadas: int) -> None:
        """Atualiza o contador de NFSe baixadas."""
        self._safe_label_update(
            self.label_nfse, 
            f"NFSe baixadas: {nfse_baixadas}"
        )

    def _safe_label_update(self, label: tk.Label, text: str) -> None:
        """Atualiza um label de forma segura, evitando erros TclError."""
        try:
            if self.winfo_exists():
                label.config(text=text)
                self.update_idletasks()
        except tk.TclError:
            # Ignora erro se o widget já foi destruído
            pass

    def finalizar(self) -> None:
        """Finaliza o popup de forma segura."""
        try:
            if hasattr(self, 'progressbar') and self.progressbar:
                self.progressbar.stop()
        except tk.TclError:
            pass
        
        try:
            if self.winfo_exists():
                self.destroy()
        except tk.TclError:
            pass

def buttons_frame(parent, buttons_config, orientacao=tk.HORIZONTAL):
    """
    Cria um frame com botões de forma genérica.
    
    Args:
        parent: Widget pai
        buttons_config: Lista de dicionários com configuração dos botões
        orientacao: Orientação do empacotamento
    
    Returns:
        tuple: (frame, dicionário com referências aos botões)
    """
    frame = tk.Frame(parent)
    frame.pack(fill=tk.X, padx=10, pady=10)
    
    buttons = {}
    pack_kwargs = {'side': tk.LEFT, 'padx': 5} if orientacao == tk.HORIZONTAL else {'side': tk.TOP, 'pady': 2}
    
    for config in buttons_config:
        btn = tk.Button(frame, **config)
        btn.pack(**pack_kwargs)
        buttons[config['text']] = btn
    
    return frame, buttons

