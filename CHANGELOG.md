---
created: 2025-10-21T14:50
updated: 2025-10-29T15:21
---
# Changelog
Todas as mudanças neste projeto serão documentadas neste arquivo.

## [1.0] - 2025-12-18
Contribuintes: Solivan A. dos Santos
### Added
- Modo de Busca por Competência ou Emissão
- Modo de Cadastro por CNPJ ou Código
- Opção de resetar os NSU de todas as empresas cadastradas
- Notas canceladas agora são zeradas no relatório excel
- Inserido atualização recursiva das planilhas em caso de alteração do VBA ou planilha mãe, em /docs/att_planilhas.py

### Fixed
- Não pula mais NFSe nas mudanças de lote
- Corrigido a perca de arquivos no reprocessamento de períodos já solicitados
- Corrigido os retornos visuais nas planilha de relatórios
- Agora permite deletar normalmente empresas que não possuem certificado cadastrado
- Janelas modais agora são fixadas e ficam em primeiro plano, movimento unificado em bloco

### Changed
- CNPJ agora aceita caracteres alfanuméricos
- Arquivos baixados começa com Prefixo definido + NSU-[n° do NSU], assim ficando mais fácil verificar a falta de arquivos durante o processo

### Removed
- Removido o campo de Úlitmo NSU Verificado

## [0.5.5] – 2025-10-29
### Fixed
- Corrigido Cod salvo como string ao invés de int

## [0.5.4] – 2025-10-21 
Contribuintes: Solivan A. dos Santos
### Added
- Agora aceita cadastro de múltiplas empresas, com controle de NSU individualmente;
- Downloads são salvos em pasta interna para serem compactados;
- Adicionado o menu 'Documentação' que possui o manual básico do programa;
- Tooltips nas configurações para maior entendimento do usuário;
- Controle de vencimento do certificado importado;
- Adicionado uma planilha .xlsm com vba que importa e gera relatório para conferência;
- Se forem baixados os .pdf, hyperlinks relativos (desde que sejam mantidos as mesmas pastas ao descompatar o .zip) que abrem a DANFSe;
- Controle de registros por competência, arquivos são baixados da competência escolhida, possui auditoria interna dos registros de NSU para isso

### Changed
- Menu inicial foi totalmente reformulado;
- Controle de NSU passou a ser por empresa cadastrada;
- Registros em .txt internos foram mudados para .json e .log;
- Arquivos agora são manejados internamente, tanto para configuração da própria execução (como certificado .pfx) ou exportação (arquivos .xml e .pdf são compactados a cada execução);
- Progesso agora é visualizado em um popup ao invés de monitoramento via log integrado

### Fixed
- API de download de .pdf foi atualizada

### Removed
- Integração do log na tela inicial
- Removido 'Autostart'
