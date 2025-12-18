from tkhtmlview import HTMLLabel
import markdown
import logging
import os
import sys
import tkinter as tk
from tkinter.scrolledtext import ScrolledText
## Módulos auxiliares
from config.config import DIRETORIOS, ROOT_DIR, Config
from ui.cad_window import CadastroUI
from ui.config_window import ConfigUI
from ui.download_window import DownloadUI
from config.config import configurar_logging
from ui.ui_basic import modal_window, back_window, centralizar

configurar_logging(nivel=logging.INFO)
logger = logging.getLogger(__name__)

try:
    from docs.version import __version__
except Exception:
    __version__ = "0.0.0"

from docs.license_text import LICENSE_TEXT

class App:
    from ui.ui_basic import _set_window_icon

    def __init__(self, root, config: Config):
        self.root = root
        self.config = config
        self.root.title(f"Download NFSe Nacional v{__version__}")
        centralizar(self.root, 500, 500)
       
        self._set_window_icon()
        
        self._create_main_interface()
        self._setup_window_references()

    def _create_main_interface(self):
        """Cria a interface principal"""
        main_frame = tk.Frame(self.root, padx=20, pady=20)
        main_frame.pack(expand=True, fill=tk.BOTH)
        
        tk.Label(
            main_frame, 
            text=f"Download NFS-e Portal Nacional\nv{__version__}", 
            font=("Arial", 16, "bold"),
            justify=tk.CENTER
        ).pack(pady=10)
        
        button_frame = tk.Frame(main_frame)
        button_frame.pack(expand=True)
        
        buttons = [
            ("Baixar NFSe", lambda: DownloadUI(self)),
            ("Cadastros", lambda: CadastroUI(self)),
            ("Configurações", lambda: ConfigUI(self)),
            ("Documetação", self.show_instructions),
        ]
        
        for text, command in buttons:
            self._create_button(button_frame, text, command).pack(pady=10)
        
        # Botão Sobre no estilo pequeno
        tk.Button(
            main_frame, text="Sobre", command=self.show_about,
            width=8, height=1, font=("Arial", 8)
        ).pack(side=tk.BOTTOM, anchor=tk.SW, padx=20, pady=10)

    def _create_button(self, parent, text, command, **kwargs):
        """Cria botões de forma padronizada"""
        style = {"width": 40, "height": 2, "font": ("Arial", 12)}
        style.update(kwargs)
        return tk.Button(parent, text=text, command=command, **style)

    def _setup_window_references(self):
        """Inicializa referências das janelas"""
        self.settings_win = self.about_win = self.instructions_win = None

    def _window_exists(self, window):
        """Verifica se uma janela ainda existe"""
        try:
            return window.winfo_exists()
        except Exception:
            return False
        
    def bring_all_to_front(self):
        """Trazer todas as janelas modais para frente junto com a principal"""
        self.root.lift()
        self.root.focus_force()
        
        # Traz todas as janelas modais para frente
        if hasattr(self.root, '_modal_windows'):
            for modal in self.root._modal_windows:
                if modal.winfo_exists():
                    try:
                        modal.lift()
                    except:
                        pass

    def show_instructions(self) -> None:
        """Exibe as instruções do arquivo instrucoes.md"""
        # Traz todas as janelas para frente primeiro
        self.bring_all_to_front()
        
        if self._window_exists(getattr(self, 'instructions_win', None)):
            self.instructions_win.lift()
            self.instructions_win.focus_set()
            return

        self.instructions_win = modal_window(self.root, "Documentação - Download NFSe Nacional", 800, 600)
        self._setup_instructions_content()
        
        # Botão Voltar no estilo pequeno
        tk.Button(
            self.instructions_win, text="Voltar",
            command=lambda: back_window(self.instructions_win, self.root),
            width=8, height=1, font=("Arial", 8)
        ).pack(side=tk.BOTTOM, anchor=tk.SW, padx=10, pady=10)

    def _setup_instructions_content(self):
        """Configura o conteúdo das instruções"""
        try:
            self._setup_html_instructions(self.instructions_win)
        except ImportError:
            self._setup_text_instructions(self.instructions_win)

    def _setup_html_instructions(self, parent):
        """Configura instruções com renderização HTML"""
        try:
            with open(DIRETORIOS['instrucoes'], "r", encoding="utf-8") as f:
                html_content = markdown.markdown(f.read())
            
            scroll_frame = tk.Frame(parent)
            scroll_frame.pack(fill=tk.BOTH, expand=True)
            
            scrollbar = tk.Scrollbar(scroll_frame)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            
            html_label = HTMLLabel(scroll_frame, html=html_content, height=20, width=70)
            html_label.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            
            scrollbar.config(command=html_label.yview)
            html_label.config(yscrollcommand=scrollbar.set)
            
        except Exception as e:
            tk.Label(parent, text=f"Erro ao carregar instruções: {str(e)}").pack(pady=20)

    def _setup_text_instructions(self, parent):
        """Configura instruções em texto simples"""
        text_area = ScrolledText(parent, wrap=tk.WORD, width=80, height=30)
        text_area.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        try:
            with open(os.path.join(ROOT_DIR, "instrucoes.md"), "r", encoding="utf-8") as f:
                text_area.insert(tk.END, f.read())
        except Exception as e:
            text_area.insert(tk.END, f"Erro ao carregar instruções: {str(e)}")
        
        text_area.config(state=tk.DISABLED)

    def show_about(self) -> None:
        """Exibe informações sobre a aplicação"""
        # Traz todas as janelas para frente primeiro
        self.bring_all_to_front()
        
        if self._window_exists(getattr(self, 'about_win', None)):
            self.about_win.lift()
            self.about_win.focus_set()
            return

        self.about_win = modal_window(self.root, "Sobre - Download NFSe Nacional", 600, 500)
        self._setup_about_content()
        
        # Botão Voltar no estilo pequeno
        tk.Button(
            self.about_win, text="Voltar",
            command=lambda: back_window(self.about_win, self.root),
            width=8, height=1, font=("Arial", 8)
        ).pack(side=tk.BOTTOM, anchor=tk.SW, padx=10, pady=10)

    def _setup_about_content(self):
        """Configura o conteúdo da janela sobre com texto justificado"""
        # Criar o widget de texto
        text = ScrolledText(self.about_win, width=80, height=25, wrap=tk.WORD)
        text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        text.tag_configure("center", justify='center')
        text.tag_configure("justify", justify='center')  
        
        about_text = (
            f"Download NFS-e Portal Nacional v{__version__}\n"
            "Autor: Renan R. Santos \nContribuidor: Solivan A. dos Santos\n\n"
            f"{LICENSE_TEXT}"
        )
        
        # Inserir o texto e configurar tags
        text.insert(tk.END, about_text)
        text.tag_add("center", "1.0", "2.end")
        text.tag_add("justify", "3.0", "end")
        text.config(state=tk.DISABLED)

if __name__ == "__main__":
    try:
        cfg = Config.load(DIRETORIOS['config_json'])
        logger.info(f"Configuração carregada: \n{cfg}")
    except Exception as e:
        tk.Tk().withdraw()
        tk.messagebox.showerror("Erro de configuração", str(e))
        sys.exit(1)

    root = tk.Tk()
    App(root, cfg)
    root.mainloop()