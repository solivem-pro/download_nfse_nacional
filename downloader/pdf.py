import logging
from typing import Optional
logger = logging.getLogger(__name__)

class NFSePDFDownloader:
    """Downloader para documentos PDF do portal nacional de NFS-e."""

    BASE_URL = "https://adn.nfse.gov.br/danfse"
    
    def __init__(self, session, timeout: int = 30):
        self.session = session
        self.timeout = timeout
        # Remove logger duplicado - usa o do módulo

    def baixar(self, chave: str, dest_path: str) -> bool:
        """
        Baixa um PDF da chave fornecida.
        
        Args:
            chave: Chave da NFS-e
            dest_path: Caminho de destino do arquivo
            
        Returns:
            True se download bem-sucedido, False caso contrário
        """
        url = f"{self.BASE_URL}/{chave}"
        
        try:
            with self.session.get(url, timeout=self.timeout) as resp:
                if resp.status_code == 200:
                    self._salvar_arquivo(dest_path, resp.content)
                    logger.info("PDF baixado com sucesso: %s", chave)
                    return True
                else:
                    logger.error("Falha ao baixar PDF %s: HTTP %s", chave, resp.status_code)
                    return False
                    
        except Exception as e:
            logger.error("Erro ao baixar PDF %s: %s", chave, str(e))
            return False

    def _salvar_arquivo(self, dest_path: str, content: bytes) -> None:
        """Salva o conteúdo no caminho especificado."""
        with open(dest_path, "wb") as f:
            f.write(content)

    def baixar_lote(self, chaves_destinos: list[tuple[str, str]]) -> tuple[int, int]:
        """
        Baixa múltiplos PDFs em lote.
        
        Args:
            chaves_destinos: Lista de tuplas (chave, caminho_destino)
            
        Returns:
            Tupla (sucessos, falhas)
        """
        sucessos = 0
        falhas = 0
        
        for chave, destino in chaves_destinos:
            if self.baixar(chave, destino):
                sucessos += 1
            else:
                falhas += 1
                
        logger.info("Lote concluído: %d sucessos, %d falhas", sucessos, falhas)
        return sucessos, falhas