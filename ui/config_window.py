import tkinter as tk
from tkinter import messagebox, ttk
from dataclasses import asdict
import logging
## Módulos auxiliares
from config.config import Config, DIRETORIOS
from ui.ui_basic import ToolTip, centralizar, back_window, modal_window
logger = logging.getLogger(__name__)
    
## ------------------------------------------------------------------------------
## Interface principal de configurações
## ------------------------------------------------------------------------------
class ConfigUI:
    """Janela de configurações modal"""
    
    TOOLTIPS = {
        "file_prefix": "Como o nome do arquivo baixado será iniciado.",
        "delay_seconds": "Tempo de download entre requisições de lotes (em segundos).\nGarantir um delay maior pode evitar bloqueios temporários.",
        "timeout": "Tempo máximo de espera por resposta do servidor (em segundos).",
        "consult_mode": "Modo de consulta por data de Competência ou Emissão. \nEm competência busca pela emissão também para evitar perca de NFSe.",
        "save_mode" : "Modo de salvamento dos cadastros, se será por código da empresa ou CNPJ",
        "download_pdf": "Se marcado, baixa os arquivos em PDF. Aumenta o tempo de processamento. \nProblemas no servidor podem ocorrer e os PDFs não serem baixados."
    }

    def __init__(self, parent):
        self.parent = parent
        self.config = parent.config
        self.vars = {}
        
        # Cria janela modal
        self.win = modal_window(parent.root, "Configurações - Download NFSe Nacional", 350, 250)
        self._create_widgets()

    def _create_widgets(self):
        """Cria todos os widgets da interface"""
        self._create_entries()
        self._create_combobox()  
        self._create_checkbox()
        self._create_tooltips()
        self._create_buttons()

    def _create_entries(self):
        """Cria os campos de entrada"""
        fields = [
            (0, "file_prefix", "Prefixo Arquivo"),
            (1, "delay_seconds", "Delay (s)"),
            (2, "timeout", "Timeout (s)")
        ]
        
        for row, key, label in fields:
            tk.Label(self.win, text=label).grid(row=row, column=0, sticky="w", padx=5, pady=5)
            var = tk.StringVar(value=str(getattr(self.config, key)))
            tk.Entry(self.win, textvariable=var, width=25).grid(row=row, column=1, padx=5, pady=5)
            self.vars[key] = var

    def _create_combobox(self):
        """Cria o combobox para consult_mode"""
        row = 3
        tk.Label(self.win, text="Modo Consulta").grid(row=row, column=0, sticky="w", padx=5, pady=5)
        
        # Criar StringVar para o combobox
        self.consult_mode_var = tk.StringVar(value=self.config.consult_mode)
        
        # Criar Combobox com as opções
        consult_mode_combo = ttk.Combobox(
            self.win, 
            textvariable=self.consult_mode_var,
            values=["Competência", "Emissão"],
            state="readonly",
            width=22
        )
        consult_mode_combo.grid(row=row, column=1, padx=5, pady=5)
        
        # Adicionar ao dicionário de variáveis
        self.vars["consult_mode"] = self.consult_mode_var

        row = 4
        tk.Label(self.win, text="Modo Cadastros").grid(row=row, column=0, sticky="w", padx=5, pady=5)
        
        # Criar StringVar para o combobox
        self.save_mode_var = tk.StringVar(value=self.config.save_mode)
        
        # Criar Combobox com as opções
        save_mode_combo = ttk.Combobox(
            self.win, 
            textvariable=self.save_mode_var,
            values=["Código", "CNPJ"],
            state="readonly",
            width=22
        )
        save_mode_combo.grid(row=row, column=1, padx=5, pady=5)
        
        # Adicionar ao dicionário de variáveis
        self.vars["save_mode"] = self.save_mode_var

    def _create_checkbox(self):
        """Cria o checkbox para download de PDF"""
        self.pdf_var = tk.BooleanVar(value=bool(self.config.download_pdf))
        chk_pdf = tk.Checkbutton(self.win, text="Baixar PDF", variable=self.pdf_var)
        chk_pdf.grid(row=5, column=1, sticky="w", padx=5, pady=5)

    def _create_tooltips(self):
        """Adiciona tooltips aos campos"""
        # Tooltips para os campos de entrada
        for row, key in enumerate(["file_prefix", "delay_seconds", "timeout"]):
            icon = tk.Label(self.win, text="❓", fg="blue", cursor="question_arrow")
            icon.grid(row=row, column=2, sticky="w", padx=2)
            ToolTip(icon, self.TOOLTIPS.get(key, ""))

        # Tooltip para consult_mode (row 3)
        icon_consult = tk.Label(self.win, text="❓", fg="blue", cursor="question_arrow")
        icon_consult.grid(row=3, column=2, sticky="w", padx=2)
        ToolTip(icon_consult, self.TOOLTIPS.get("consult_mode", ""))

        # Tooltip para save_mode (row 4)
        icon_save = tk.Label(self.win, text="❓", fg="blue", cursor="question_arrow")
        icon_save.grid(row=4, column=2, sticky="w", padx=2)
        ToolTip(icon_save, self.TOOLTIPS.get("save_mode", ""))

        # Tooltip para download_pdf (row 5)
        icon_pdf = tk.Label(self.win, text="❓", fg="blue", cursor="question_arrow")
        icon_pdf.grid(row=5, column=2, sticky="w", padx=2)
        ToolTip(icon_pdf, self.TOOLTIPS.get("download_pdf", ""))

    def _create_buttons(self):
        """Cria os botões de ação"""
        frame_buttons = tk.Frame(self.win)
        frame_buttons.grid(row=6, column=0, columnspan=3, padx=60, pady=15)

        tk.Button(frame_buttons, text="Salvar", width=12, command=self._save).grid(row=0, column=0, padx=10)
        tk.Button(frame_buttons, text="Cancelar", width=12, command=self._on_close).grid(row=0, column=1, padx=10)

    def _save(self):
        """Salva as configurações"""
        new_data = asdict(self.config)
        
        try:
            for key, var in self.vars.items():
                if key in ("delay_seconds", "timeout"):
                    new_data[key] = float(var.get())
                else:
                    new_data[key] = var.get()
                logger.info(f"Configuração de {key} definida {new_data[key]}")
                    
            new_data["download_pdf"] = self.pdf_var.get()
            
            # Atualiza a configuração no parent (janela principal)
            self.parent.config = Config(**new_data)
            
            # Salva no arquivo
            self.parent.config.save(DIRETORIOS['config_json'])
            
            messagebox.showinfo("Configurações", "Configurações salvas com sucesso!")

            self._on_close()
            
        except ValueError:
            messagebox.showerror("Erro", "Valores numéricos inválidos!")
        
    def _on_close(self):
        """Handler para fechamento da janela"""
        back_window(self.win, self.parent.root)

def ler_config() -> Config:
    return Config.load(DIRETORIOS['config_json'])