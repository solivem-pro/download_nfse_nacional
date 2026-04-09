---
created: 2025-10-21T14:50
updated: 2026-04-09T11:08
---
# Changelog

Todas as mudancas relevantes deste projeto serao documentadas neste arquivo.

## [Unreleased] - 2026-04-08
Contribuinte: Thiago V. M. dos Santos

### Added
- Shell principal em `PySide6` como interface padrao do projeto.
- Persistencia de cadastros em `SQLite`, com base `cadastros.db`.
- Importacao de empresas por planilha `.xlsx` na tela de `Cadastros`.
- Guia operacional atualizado para a primeira execucao da interface Qt.
- Download paralelo de DANFSe por chave de acesso.
- Consulta paralela de eventos por chave da NFS-e, com salvamento dos XMLs em `EVENTOS`.
- Identificacao de cancelamento por evento para separar documentos cancelados em subpasta dedicada.
- Resumo do download com estatisticas separadas para PDFs, eventos e cancelamentos.
- Geracao nativa de relatorio PDF a partir dos XMLs baixados, sem macro do Excel.

### Changed
- Documentacao principal atualizada para refletir o launcher `download_nfse_qt.py`.
- Build Windows atualizado para gerar executavel `onefile`, com arquivos internos em `%AppData%` e downloads na area de trabalho do usuario.
- Instalador ajustado para o ambiente atual, sem dependencia de `tkinter`.
- Fluxo de cadastros migrado do JSON legado para `SQLite`, com compatibilidade de leitura e importacao inicial.
- Preparacao automatica da estrutura fisica da empresa no primeiro uso, sem depender de planilha modelo.
- Download de PDFs e eventos passou a usar `retry`, `backoff` e `delay` proprio para reduzir falhas por rate limit e instabilidade.
- Classificacao de XMLs de evento passou a priorizar `infEvento/pedRegEvento`, evitando misturar evento com NFSe normal.
- Estrutura de `packs` reorganizada para `Empresa/Competencia (MM-AAAA)/PRESTADOS|TOMADOS|EVENTOS`.
- NFS-e passam a ser salvas com o formato `<Numero da NFSe> - <Chave de acesso>`.
- NFS-e com retencao passam a receber o sufixo `- Retido`.
- Arquivo de controle `nsu_competencia.json` passa a ficar na pasta raiz da empresa.
- Pos-processamento e compactacao passam a considerar a competencia atual, sem misturar periodos diferentes no mesmo pacote.
- Dependencias e documentacao passam a refletir o fluxo nativo em PDF, sem `pywin32` no caminho principal.

### Removed
- Launcher antigo em Tk e scripts da antiga pasta `ui/`.
- Dependencias e referencias legadas de `tkhtmlview`, `markdown` e `plyer`.
- Dependencia ativa de planilhas `.xlsm` e macro do Excel no fluxo principal de download.

## [1.0] - 2025-12-18
Contribuinte: Solivan A. dos Santos

### Added
- Modo de busca por competencia ou emissao.
- Modo de cadastro por CNPJ ou codigo.
- Opcao de resetar os NSUs de todas as empresas cadastradas.
- Notas canceladas agora sao zeradas no relatorio Excel.
- Atualizacao recursiva das planilhas em caso de alteracao do VBA ou planilha mae, em `config/att_planilhas.py`.

### Fixed
- Nao pula mais NFS-e nas mudancas de lote.
- Corrigida a perda de arquivos no reprocessamento de periodos ja solicitados.
- Corrigidos os retornos visuais nas planilhas de relatorio.
- Agora permite excluir normalmente empresas que nao possuem certificado cadastrado.
- Janelas modais passaram a ficar fixadas e em primeiro plano.

### Changed
- CNPJ passou a aceitar caracteres alfanumericos.
- Arquivos baixados passaram a iniciar com o prefixo definido mais `NSU-<numero>`, facilitando a conferencia.

### Removed
- Campo de ultimo NSU verificado.

## [0.5.5] - 2025-10-29

### Fixed
- Corrigido codigo salvo como string em vez de inteiro.

## [0.5.4] - 2025-10-21
Contribuinte: Solivan A. dos Santos

### Added
- Cadastro de multiplas empresas com controle de NSU individual.
- Downloads salvos em pasta interna para compactacao posterior.
- Menu de documentacao com manual basico do programa.
- Tooltips nas configuracoes para maior entendimento do usuario.
- Controle de vencimento do certificado importado.
- Planilha `.xlsm` com VBA para importacao e relatorio de conferencia.
- Hyperlinks relativos para abrir a DANFSe quando os PDFs sao baixados.
- Controle de registros por competencia com auditoria interna dos NSUs.

### Changed
- Menu inicial reformulado.
- Controle de NSU passou a ser por empresa cadastrada.
- Registros internos em `.txt` migraram para `.json` e `.log`.
- Arquivos passaram a ser manejados internamente para configuracao e exportacao.
- Progresso passou a ser visualizado em popup, substituindo o log integrado em tela.

### Fixed
- API de download de PDF atualizada.

### Removed
- Integracao do log na tela inicial.
- Removido `Autostart`.
