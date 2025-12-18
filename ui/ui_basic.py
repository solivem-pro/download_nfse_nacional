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

def centralizar_em_parent(parent, janela, largura: int, altura: int) -> None:
    """Centraliza uma janela em relação à janela pai (parent)."""
    # CRÍTICO: Força atualização múltiplas vezes para garantir coordenadas corretas
    for _ in range(3):
        parent.update_idletasks()
        janela.update_idletasks()
    
    # Obtém posição e dimensões da parent
    parent_x = parent.winfo_x()
    parent_y = parent.winfo_y()
    parent_width = parent.winfo_width()
    parent_height = parent.winfo_height()
    
    # Calcula posição para centralizar na parent
    x = parent_x + (parent_width - largura) // 2
    y = parent_y + (parent_height - altura) // 2
    
    # Define a geometria SEM validação de tela - deixa o SO gerenciar
    janela.geometry(f"{largura}x{altura}+{x}+{y}")
    janela.resizable(False, False)

def modal_window(parent, titulo, largura, altura):
    """Cria uma janela modal padronizada que acompanha a parent"""
    win = tk.Toplevel(parent)
    win.title(titulo)
    
    # Tenta definir o mesmo ícone da janela principal
    try:
        from config.config import DIRETORIOS
        if DIRETORIOS['icone'].exists():
            if str(DIRETORIOS['icone']).endswith('.ico'):
                win.iconbitmap(str(DIRETORIOS['icone']))
    except Exception:
        pass
    
    # Centraliza em relação à parent
    centralizar_em_parent(parent, win, largura, altura)
    
    win.transient(parent)
    win.grab_set()
    win.focus_set()
    
    # Torna a janela modal sempre visível sobre a parent
    win.lift()
    win.attributes('-topmost', True)
    win.attributes('-topmost', False)
    
    # Encontra a janela raiz (janela principal)
    root_window = parent
    while hasattr(root_window, '_parent_window') and root_window._parent_window is not None:
        root_window = root_window._parent_window
    
    # CORREÇÃO: Inicializa estrutura de controle de modais na raiz SEMPRE
    if not hasattr(root_window, '_modal_windows'):
        root_window._modal_windows = []
        root_window._is_moving = False
    
    # Marca a hierarquia
    win._parent_window = parent
    win._root_window = root_window
    
    # Adiciona à lista de modais da raiz
    root_window._modal_windows.append(win)
    
    # Desabilita apenas a janela pai direta, não a raiz
    parent.attributes('-disabled', True)
    
    # Calcula offset em relação à parent
    win.update_idletasks()
    parent.update_idletasks()
    win._parent_offset = (win.winfo_x() - parent.winfo_x(), 
                         win.winfo_y() - parent.winfo_y())
    win._is_moving = False
    win._last_position = (win.winfo_x(), win.winfo_y())
    
    # Sistema de movimento sincronizado hierárquico
    def _sync_windows(source_window):
        """Sincroniza posição baseado na hierarquia"""
        # Se a fonte for uma modal, ela move sua parent e siblings
        if source_window != root_window:
            # É uma modal
            if getattr(source_window, '_is_moving', False):
                return
            source_window._is_moving = True
            
            try:
                current_x = source_window.winfo_x()
                current_y = source_window.winfo_y()
                
                if not hasattr(source_window, '_last_position'):
                    source_window._last_position = (current_x, current_y)
                    return
                    
                last_x, last_y = source_window._last_position
                
                delta_x = current_x - last_x
                delta_y = current_y - last_y
                
                if abs(delta_x) < 3 and abs(delta_y) < 3:
                    return
                
                # Move a parent direta
                parent_win = source_window._parent_window
                if parent_win and parent_win.winfo_exists():
                    parent_x = parent_win.winfo_x()
                    parent_y = parent_win.winfo_y()
                    new_parent_x = parent_x + delta_x
                    new_parent_y = parent_y + delta_y
                    
                    parent_w = parent_win.winfo_width()
                    parent_h = parent_win.winfo_height()
                    parent_win.geometry(f"{parent_w}x{parent_h}+{new_parent_x}+{new_parent_y}")
                    
                    # Atualiza todas as modais da mesma parent
                    for modal in root_window._modal_windows[:]:
                        if not modal.winfo_exists():
                            root_window._modal_windows.remove(modal)
                            continue
                        
                        if modal._parent_window == parent_win and modal != source_window:
                            if hasattr(modal, '_parent_offset'):
                                offset_x, offset_y = modal._parent_offset
                                modal_x = new_parent_x + offset_x
                                modal_y = new_parent_y + offset_y
                                
                                modal_w = modal.winfo_width()
                                modal_h = modal.winfo_height()
                                modal.geometry(f"{modal_w}x{modal_h}+{modal_x}+{modal_y}")
                                if hasattr(modal, '_last_position'):
                                    modal._last_position = (modal_x, modal_y)
                
                # Atualiza offset
                if parent_win and parent_win.winfo_exists():
                    source_window._parent_offset = (current_x - parent_win.winfo_x(), 
                                                  current_y - parent_win.winfo_y())
                source_window._last_position = (current_x, current_y)
                
            finally:
                source_window._is_moving = False
                
        else:
            # A raiz foi movida
            if getattr(root_window, '_is_moving', False):
                return
            root_window._is_moving = True
            
            try:
                root_x = root_window.winfo_x()
                root_y = root_window.winfo_y()
                
                # Move todas as modais da raiz
                for modal in root_window._modal_windows[:]:
                    if not modal.winfo_exists():
                        root_window._modal_windows.remove(modal)
                        continue
                    
                    # Encontra a chain de parents até a raiz
                    current_parent = modal._parent_window
                    offset_x = modal.winfo_x() - current_parent.winfo_x()
                    offset_y = modal.winfo_y() - current_parent.winfo_y()
                    
                    # Calcula nova posição baseada na posição atual da parent
                    new_x = current_parent.winfo_x() + offset_x
                    new_y = current_parent.winfo_y() + offset_y
                    
                    modal_w = modal.winfo_width()
                    modal_h = modal.winfo_height()
                    modal.geometry(f"{modal_w}x{modal_h}+{new_x}+{new_y}")
                    if hasattr(modal, '_last_position'):
                        modal._last_position = (new_x, new_y)
                        
            finally:
                root_window._is_moving = False
    
    # Handler de movimento
    def _on_configure(event):
        if event.widget == win:
            _sync_windows(win)
        elif event.widget == root_window:
            _sync_windows(root_window)
        elif hasattr(event.widget, '_parent_window'):
            _sync_windows(event.widget)
    
    # Vincula eventos apenas se não estiverem vinculados
    if not hasattr(root_window, '_configure_bound'):
        root_window.bind('<Configure>', _on_configure)
        root_window._configure_bound = True
    
    win.bind('<Configure>', _on_configure)
    
    # NOVA FUNÇÃO: Reativação hierárquica completa
    def reativar_hierarquicamente(janela_atual):
        """Reativa todas as janelas na hierarquia até a raiz"""
        # Encontra a raiz
        root_win = janela_atual
        while hasattr(root_win, '_parent_window') and root_win._parent_window is not None:
            root_win = root_win._parent_window
        
        # Reativa todas as janelas da raiz para baixo
        def _reativar_recursiva(j):
            if j and j.winfo_exists():
                try:
                    j.attributes('-disabled', False)
                    j.focus_set()
                except:
                    pass
                # Reativa filhos
                if hasattr(j, '_modal_windows'):
                    for modal in j._modal_windows:
                        if modal.winfo_exists():
                            _reativar_recursiva(modal)
        
        _reativar_recursiva(root_win)
        return root_win
    
    # Handler de fechamento com reativação hierárquica
    def on_closing():
        """Handler para fechamento da janela modal"""
        try:
            # Fecha todas as modais filhas primeiro
            child_modals = []
            for modal in root_window._modal_windows[:]:
                if (hasattr(modal, '_parent_window') and 
                    modal._parent_window == win and 
                    modal.winfo_exists()):
                    child_modals.append(modal)
            
            for child_modal in child_modals:
                try:
                    child_modal._on_closing()
                except:
                    try:
                        child_modal.destroy()
                    except:
                        pass
            
            # Remove da lista da raiz ANTES de verificar outras modais
            if hasattr(root_window, '_modal_windows') and win in root_window._modal_windows:
                root_window._modal_windows.remove(win)
            
            # Limpa atributos
            for attr in ['_parent_window', '_root_window', '_parent_offset', 
                        '_is_moving', '_last_position']:
                if hasattr(win, attr):
                    delattr(win, attr)
            
            # Libera grab antes de destruir
            try:
                win.grab_release()
            except:
                pass
            
            # NOVA LÓGICA: Reativação hierárquica
            if hasattr(root_window, '_modal_windows'):
                # Remove esta modal da lista
                if win in root_window._modal_windows:
                    root_window._modal_windows.remove(win)
                
                # Se não há mais modais, reativa a hierarquia completa
                if len(root_window._modal_windows) == 0:
                    root_win = reativar_hierarquicamente(parent)
                    # Garante que a raiz fique em primeiro plano
                    if root_win and root_win.winfo_exists():
                        try:
                            root_win.lift()
                            root_win.focus_force()
                        except:
                            pass
                else:
                    # Ainda há modais, reativa apenas a parent direta
                    # mas apenas se não for parent de outras modais
                    parent_modal_count = 0
                    for modal in root_window._modal_windows[:]:  # Usar cópia
                        if (hasattr(modal, '_parent_window') and 
                            modal._parent_window == parent and 
                            modal.winfo_exists()):
                            parent_modal_count += 1
                    
                    if parent_modal_count == 0:
                        try:
                            parent.attributes('-disabled', False)
                            parent.focus_set()
                            parent.lift()
                        except:
                            pass
            
            # Destrói a modal
            win.destroy()
            
            # Limpa bindings da raiz se não há mais modais
            if (hasattr(root_window, '_modal_windows') and 
                len(root_window._modal_windows) == 0 and
                hasattr(root_window, '_configure_bound')):
                try:
                    root_window.unbind('<Configure>')
                except:
                    pass
                del root_window._configure_bound
                    
        except Exception as e:
            logger.error(f"Erro ao fechar modal: {e}")
            # EMERGENCIAL: Força reativação da raiz
            try:
                root_win = reativar_hierarquicamente(parent)
                if root_win and root_win.winfo_exists():
                    root_win.attributes('-disabled', False)
                    root_win.focus_force()
            except:
                pass
    
    win.protocol("WM_DELETE_WINDOW", on_closing)
    win._on_closing = on_closing
    
    # Fecha modais quando a raiz fecha
    def _on_root_close(event=None):
        if hasattr(root_window, '_modal_windows'):
            for modal in root_window._modal_windows[:]:
                if modal.winfo_exists():
                    try:
                        modal._on_closing()
                    except:
                        try:
                            modal.destroy()
                        except:
                            pass
    
    if not hasattr(root_window, '_close_bound'):
        root_window.bind('<Destroy>', _on_root_close, add='+')
        root_window._close_bound = True
    
    # Mantém a hierarquia correta de sobreposição
    def maintain_hierarchy(event=None):
        if not win.winfo_exists():
            return
        
        # Mantém esta modal acima de sua parent
        parent_win = win._parent_window if hasattr(win, '_parent_window') else None
        if parent_win and parent_win.winfo_exists():
            win.lift(parent_win)
        
        # Mantém as modais filhas acima desta
        for modal in root_window._modal_windows[:]:
            if (hasattr(modal, '_parent_window') and 
                modal._parent_window == win and 
                modal.winfo_exists()):
                modal.lift(win)
    
    win.bind('<Map>', maintain_hierarchy)
    win.bind('<FocusIn>', maintain_hierarchy)
    parent.bind('<FocusIn>', lambda e: maintain_hierarchy(), add='+')
    
    return win

def back_window(win, root):
    """Fecha a janela modal e retorna o foco para a janela principal"""
    if hasattr(win, '_on_closing'):
        win._on_closing()
    else:
        try:
            # Tenta usar a lógica padrão de fechamento
            if hasattr(win, '_parent_window') and win._parent_window:
                parent = win._parent_window
                # Remove da lista da raiz
                root_window = getattr(win, '_root_window', root)
                if hasattr(root_window, '_modal_windows') and win in root_window._modal_windows:
                    root_window._modal_windows.remove(win)
                
                # Reativa a parent se não houver outras modais
                if parent and parent.winfo_exists():
                    parent.attributes('-disabled', False)
                    parent.focus_set()
                    parent.lift()
            
            win.destroy()
            try:
                root.grab_release()
            except:
                pass
            root.lift()
            root.focus_set()
        except:
            pass

def _set_window_icon(self):
    """Define o ícone da janela principal"""
    try:
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
        columns_config: Lista de tuplas (id, texto, largura, anchor) ou (id, texto, largura, anchor, tipo_ordenacao)
        height: Altura do treeview
        show: Configuração de exibição
    
    Returns:
        tuple: (treeview, scrollbar, frame_tree)
    """
    frame_tree = tk.Frame(parent)
    frame_tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
    # Processar columns_config para compatibilidade (4 ou 5 elementos)
    expanded_config = []
    for col in columns_config:
        if len(col) == 4:
            # Configuração antiga: adicionar tipo de ordenação padrão
            col_id, heading, width, anchor = col
            expanded_config.append((col_id, heading, width, anchor, 'string'))
        elif len(col) == 5:
            # Configuração nova: já inclui tipo de ordenação
            expanded_config.append(col)
        else:
            raise ValueError(f"Tupla inválida em columns_config: {col}")
    
    columns = [col[0] for col in expanded_config]
    tree = ttk.Treeview(frame_tree, columns=columns, show=show, height=height)
    
    # Configurar cabeçalhos e colunas COM ORDENAÇÃO
    for col_id, heading, width, anchor, sort_type in expanded_config:
        # Criar uma closure para capturar os valores corretos
        def make_sort_func(c=col_id, t=sort_type):
            return lambda: sort_treeview_column(tree, c, t)
        
        tree.heading(col_id, text=heading, anchor=anchor,
                    command=make_sort_func())  # Configurar comando de clique
        tree.column(col_id, width=width, anchor=anchor)
    
    # Scrollbar
    scrollbar = ttk.Scrollbar(frame_tree, orient=tk.VERTICAL, command=tree.yview)
    tree.configure(yscrollcommand=scrollbar.set)
    
    tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    # Inicializar estado de ordenação
    tree.sort_states = {}
    for col_id in columns:
        tree.sort_states[col_id] = 'asc'
    
    return tree, scrollbar, frame_tree

def sort_treeview_column(tree, col, sort_type='string'):
    """
    Ordena o conteúdo do treeview por coluna.
    
    Args:
        tree: Treeview widget
        col: ID da coluna a ordenar
        sort_type: Tipo de ordenação ('string', 'int', 'float', 'date_dd_mm_yyyy')
    """
    # Alternar estado de ordenação
    if tree.sort_states[col] == 'asc':
        reverse_flag = False
        tree.sort_states[col] = 'desc'
    else:
        reverse_flag = True
        tree.sort_states[col] = 'asc'
    
    # Obter todos os itens
    items = [(tree.set(k, col), k) for k in tree.get_children('')]
    
    # Funções de conversão baseadas no tipo
    def convert_string(val):
        return str(val).lower() if val else ''
    
    def convert_int(val):
        try:
            # Remove separadores de milhar e converte
            clean_val = str(val).replace('.', '').replace(',', '').strip()
            return int(clean_val) if clean_val else 0
        except:
            return 0
    
    def convert_float(val):
        try:
            # Remove separadores de milhar e converte
            clean_val = str(val).replace('.', '').replace(',', '.').strip()
            return float(clean_val) if clean_val else 0.0
        except:
            return 0.0
    
    def convert_date(val):
        try:
            # Converte data no formato dd/mm/yyyy
            if val:
                day, month, year = map(int, str(val).split('/'))
                return year * 10000 + month * 100 + day
            return 0
        except:
            return 0
    
    # Selecionar função de conversão baseada no tipo
    converters = {
        'string': convert_string,
        'int': convert_int,
        'float': convert_float,
        'date_dd_mm_yyyy': convert_date
    }
    
    converter = converters.get(sort_type, convert_string)
    
    # Ordenar os itens
    items.sort(key=lambda x: converter(x[0]), reverse=reverse_flag)
    
    # Reorganizar itens no treeview
    for index, (_, k) in enumerate(items):
        tree.move(k, '', index)
    
    # Atualizar seta de indicação de ordenação
    update_sort_indicator(tree, col, reverse_flag)

def update_sort_indicator(tree, col, reverse):
    """
    Atualiza o cabeçalho da coluna para mostrar a direção da ordenação.
    """
    for column in tree['columns']:
        current_text = tree.heading(column)['text']
        # Remove setas existentes
        if current_text.startswith('▲ ') or current_text.startswith('▼ '):
            current_text = current_text[2:]
        
        if column == col:
            indicator = '▼ ' if reverse else '▲ '
            tree.heading(column, text=indicator + current_text)
        else:
            tree.heading(column, text=current_text)
            tree.sort_states[column] = 'asc' 

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
        
        icone_path = Path(DIRETORIOS['icone'])
        app_icon = str(icone_path) if icone_path.exists() else None
        
        notification.notify(
            title="NFS-e Nacional",
            message=mensagem,
            app_name="Download NFS-e Nacional",
            app_icon=app_icon,
            timeout=5
        )
    except Exception as e:
        print(f"Notificação falhou: {e}")

## Popup de processamento
class PopupProcessamento:
    """Popup modal para exibir progresso de operações."""

    def __init__(self, parent, titulo: str = "Processando...", texto: str = "Aguarde o processamento"):
        self.parent = parent
        self.win = modal_window(parent, titulo, 400, 180)
        self._setup_ui(texto)
        self.cancelado = False
        
        # Configurações adicionais específicas do popup
        self.win.resizable(False, False)
        
        # Guarda o handler original para usar mais tarde
        self._original_on_closing = self.win._on_closing
        
        # Handler de fechamento personalizado para popup de processamento
        def popup_on_closing():
            """Handler de fechamento específico para popup de processamento"""
            try:
                # Para a barra de progresso
                if hasattr(self, 'progressbar') and self.progressbar:
                    self.progressbar.stop()
                
                # Remove a modal da lista da raiz
                root_window = getattr(self.win, '_root_window', None)
                if root_window and hasattr(root_window, '_modal_windows'):
                    if self.win in root_window._modal_windows:
                        root_window._modal_windows.remove(self.win)
                
                # Libera o grab ANTES de reativar a parent
                try:
                    self.win.grab_release()
                except:
                    pass
                
                # Reativa a parent direta (que é a modal win1)
                parent_direct = getattr(self.win, '_parent_window', None)
                if parent_direct and parent_direct.winfo_exists():
                    # Verifica se há outras modais abertas pela mesma parent (excluindo esta)
                    other_modals = []
                    if root_window and hasattr(root_window, '_modal_windows'):
                        for modal in root_window._modal_windows[:]:  # Usar cópia
                            if (modal.winfo_exists() and
                                hasattr(modal, '_parent_window') and 
                                modal._parent_window == parent_direct and
                                modal != self.win):  # Exclui o próprio popup
                                other_modals.append(modal)
                    
                    # Se não há outras modais, reativa a parent
                    if len(other_modals) == 0:
                        try:
                            parent_direct.attributes('-disabled', False)
                            parent_direct.focus_set()
                            parent_direct.lift()
                        except:
                            pass
                
                # Destrói a janela
                if self.win.winfo_exists():
                    self.win.destroy()
                    
            except Exception as e:
                logger.error(f"Erro ao fechar popup de processamento: {e}")
                # Garante que a parent seja reativada mesmo em caso de erro
                try:
                    parent_direct = getattr(self.win, '_parent_window', None)
                    if parent_direct and parent_direct.winfo_exists():
                        parent_direct.attributes('-disabled', False)
                except:
                    pass
        
        # Substitui o handler padrão pelo nosso
        self.win.protocol("WM_DELETE_WINDOW", popup_on_closing)
        self.win._on_closing = popup_on_closing

    def winfo_exists(self):
        """Método de compatibilidade para verificar se a janela existe."""
        return self.win.winfo_exists() if hasattr(self, 'win') and self.win else False

    def _setup_ui(self, texto: str) -> None:
        """Configura a interface do usuário do popup."""
        frame_principal = tk.Frame(self.win, padx=20, pady=20)
        frame_principal.pack(fill=tk.BOTH, expand=True)
        
        tk.Label(
            frame_principal, 
            text=texto,
            font=("Arial", 10)
        ).pack(pady=(0, 15))
        
        self.label_contador = tk.Label(
            frame_principal,
            text="Processando: 0/0 - NSU: 0",
            font=("Arial", 9),
            fg="blue"
        )
        self.label_contador.pack(pady=(0, 5))
        
        self.label_nfse = tk.Label(
            frame_principal,
            text="NFSe baixadas: 0",
            font=("Arial", 9),
            fg="green"
        )
        self.label_nfse.pack(pady=(0, 10))
        
        self.progressbar = ttk.Progressbar(
            frame_principal,
            mode='indeterminate',
            length=350
        )
        self.progressbar.pack(pady=(0, 15))
        self.progressbar.start(10)

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
            if self.winfo_exists():  # Usa o método winfo_exists da classe
                label.config(text=text)
                self.win.update_idletasks()
        except tk.TclError:
            pass

    def finalizar(self) -> None:
        """Finaliza o popup de forma segura."""
        try:
            if hasattr(self, 'progressbar') and self.progressbar:
                self.progressbar.stop()
        except tk.TclError:
            pass
        
        # Chama o handler de fechamento personalizado
        if hasattr(self.win, '_on_closing'):
            self.win._on_closing()
        else:
            try:
                if self.win.winfo_exists():
                    self.win.destroy()
            except tk.TclError:
                pass

    # Método para compatibilidade com código existente
    def destroy(self):
        """Método de compatibilidade para destruir a janela."""
        self.finalizar()

    # Método para compatibilidade com código existente (se usado)
    def update_idletasks(self):
        """Método de compatibilidade para update_idletasks."""
        if self.winfo_exists():
            self.win.update_idletasks()

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