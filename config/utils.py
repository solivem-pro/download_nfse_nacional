import tkinter as tk
from typing import Optional, Union

## ----------------------------------------------------------------------------------------------
## Validadores de registros
## ----------------------------------------------------------------------------------------------
def enter_next_input(campos, funcao_final=None):
    """
    Configura navegação por Enter entre campos.
    
    Args:
        campos: Lista de campos Entry
        funcao_final: Função a ser chamada no último campo
    """
    for i, campo in enumerate(campos):
        proximo = campos[i + 1] if i < len(campos) - 1 else None
        
        if proximo:
            campo.bind("<KeyPress-Return>", lambda e, prox=proximo: prox.focus_set())
        elif funcao_final:
            campo.bind("<KeyPress-Return>", lambda e: funcao_final())

## ----------------------------------------------------------------------------------------------
## Vlidação e formatação de CNPJ
## ----------------------------------------------------------------------------------------------
def limpar_cnpj(cnpj: str) -> str:
    """Remove formatação de um CNPJ, mantendo zeros à esquerda."""
    if not cnpj:
        return ""
    return "".join(filter(str.isalnum, str(cnpj)))

def validar_cnpj(cnpj: str) -> Optional[str]:
    """Valida e retorna CNPJ limpo se for válido."""
    cnpj_limpo = limpar_cnpj(cnpj)
    return cnpj_limpo if len(cnpj_limpo) == 14 else None

def formatar_cnpj(cnpj: str) -> str:
    """Formata um CNPJ para o padrão XX.XXX.XXX/XXXX-XX."""
    cnpj_limpo = limpar_cnpj(cnpj)
    
    if len(cnpj_limpo) != 14:
        return cnpj_limpo  # Retornar sem formatação se inválido
    
    return f"{cnpj_limpo[:2]}.{cnpj_limpo[2:5]}.{cnpj_limpo[5:8]}/{cnpj_limpo[8:12]}-{cnpj_limpo[12:14]}"

def formatar_cnpj_digitacao(event: tk.Event) -> None:
    """Event handler para formatação automática de CNPJ durante digitação."""
    widget = event.widget
    texto_original = widget.get()
    cursor_pos_original = widget.index(tk.INSERT)

    # Remove tudo que não é dígito e limita a 14 caracteres
    cnpj_limpo = "".join(filter(str.isalnum, texto_original))[:14]
    
    # Aplica formatação progressiva
    formatos = [
        (2, "{}"),
        (5, "{}.{}"),
        (8, "{}.{}.{}"),
        (12, "{}.{}.{}/{}"),
        (14, "{}.{}.{}/{}-{}")
    ]
    
    for limite, formato in formatos:
        if len(cnpj_limpo) <= limite:
            partes = [cnpj_limpo[:2]]
            if limite >= 5: partes.append(cnpj_limpo[2:5])
            if limite >= 8: partes.append(cnpj_limpo[5:8])
            if limite >= 12: partes.append(cnpj_limpo[8:12])
            if limite >= 14: partes.append(cnpj_limpo[12:14])
            
            cnpj_formatado = formato.format(*partes)
            break
    else:
        cnpj_formatado = cnpj_limpo

    # Atualiza o widget
    widget.delete(0, tk.END)
    widget.insert(0, cnpj_formatado)
    
    # Reposiciona o cursor
    digitos_ate_cursor = sum(1 for ch in texto_original[:cursor_pos_original] if ch.isalnum())
    novo_cursor = 0
    digitos_encontrados = 0
    
    for char in cnpj_formatado:
        if digitos_encontrados >= digitos_ate_cursor:
            break
        if char.isalnum():
            digitos_encontrados += 1
        novo_cursor += 1
    
    widget.icursor(novo_cursor)

## ----------------------------------------------------------------------------------------------
## Validação e formatação de números
## ----------------------------------------------------------------------------------------------

def limpar_numero(valor: Union[str, int]) -> int:
    """Remove pontos, vírgulas e espaços, retornando um número inteiro."""
    if isinstance(valor, int):
        return valor
    valor_limpo = "".join(ch for ch in str(valor) if ch.isdigit())
    return int(valor_limpo or "0")

def formatar_milhar(event: tk.Event) -> None:
    widget = event.widget
    texto_original = widget.get()
    cursor_pos_original = widget.index(tk.INSERT)

    # Remove caracteres não numéricos exceto vírgula
    texto_limpo = "".join(ch for ch in texto_original if ch.isdigit() or ch == ",")
    
    # Mantém apenas a primeira vírgula para decimais
    if texto_limpo.count(",") > 1:
        partes = texto_limpo.split(",", 1)
        texto_limpo = partes[0] + "," + "".join(partes[1].replace(",", ""))

    # Se estiver vazio, não formata (mantém vazio)
    if not texto_limpo:
        return

    try:
        # Formata a parte inteira
        if "," in texto_limpo:
            inteiro_str, decimal_str = texto_limpo.split(",", 1)
            # Só converte se não estiver vazio
            inteiro = int(inteiro_str) if inteiro_str else 0
            texto_formatado = f"{inteiro:,}".replace(",", ".") + f",{decimal_str}"
        else:
            inteiro = int(texto_limpo)
            texto_formatado = f"{inteiro:,}".replace(",", ".")
    except ValueError:
        texto_formatado = texto_limpo

    # Atualiza o widget
    widget.delete(0, tk.END)
    widget.insert(0, texto_formatado)
    
    # Reposiciona o cursor
    chars_validos_ate_cursor = sum(
        1 for ch in texto_original[:cursor_pos_original] 
        if ch.isdigit() or ch == ","
    )
    
    novo_cursor = 0
    chars_encontrados = 0
    
    for char in texto_formatado:
        if chars_encontrados >= chars_validos_ate_cursor:
            break
        if char.isdigit() or char == ",":
            chars_encontrados += 1
        novo_cursor += 1
    
    widget.icursor(novo_cursor)