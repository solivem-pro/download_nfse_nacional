from typing import Optional, Union


def limpar_cnpj(cnpj: str) -> str:
    """Remove a formatacao de um CNPJ preservando zeros a esquerda."""
    if not cnpj:
        return ""
    return "".join(filter(str.isalnum, str(cnpj)))


def validar_cnpj(cnpj: str) -> Optional[str]:
    """Retorna o CNPJ limpo quando o valor possui 14 caracteres."""
    cnpj_limpo = limpar_cnpj(cnpj)
    return cnpj_limpo if len(cnpj_limpo) == 14 else None


def formatar_cnpj(cnpj: str) -> str:
    """Formata um CNPJ no padrao XX.XXX.XXX/XXXX-XX."""
    cnpj_limpo = limpar_cnpj(cnpj)
    if len(cnpj_limpo) != 14:
        return cnpj_limpo
    return f"{cnpj_limpo[:2]}.{cnpj_limpo[2:5]}.{cnpj_limpo[5:8]}/{cnpj_limpo[8:12]}-{cnpj_limpo[12:14]}"


def limpar_numero(valor: Union[str, int]) -> int:
    """Remove separadores e espacos, retornando um inteiro."""
    if isinstance(valor, int):
        return valor
    valor_limpo = "".join(ch for ch in str(valor) if ch.isdigit())
    return int(valor_limpo or "0")
