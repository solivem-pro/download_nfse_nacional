# Download NFS-e Portal Nacional

Aplicacao desktop para consulta, cadastro e download de NFS-e no Portal Nacional.

## Estado atual

- Interface principal em `PySide6`
- Entrada oficial da aplicacao em `download_nfse_qt.py`
- Build Windows em `build_exe_windows.bat`
- Backend de download, configuracao e cadastros mantido no projeto atual

## Requisitos

- Windows
- Python com suporte as dependencias de `requirements.txt`

## Instalacao

Instale as dependencias do projeto:

```powershell
pip install -r requirements.txt
```

## Como abrir

Execute a interface principal:

```powershell
python download_nfse_qt.py
```

## Build Windows

Para gerar o executavel, execute:

```powershell
build_exe_windows.bat
```

O script de build:

- instala ou atualiza as dependencias de `requirements.txt`
- instala `PyInstaller` no ambiente atual
- limpa artefatos antigos de `build/` e `dist/`
- gera o executavel unico em `dist/download_nfse.exe`

Nesse fluxo, nao e necessario executar `1.instalador.py` para empacotar o projeto.

## Menus principais

### Inicio

- resumo rapido das empresas cadastradas
- atalhos para `Download`, `Cadastros`, `Configuracoes` e `Documentacao`

### Download

- selecao de empresas por tabela
- filtros de `Ano`, `Mes` e `Status`
- execucao do fluxo real de download
- geracao automatica do relatorio PDF por competencia
- exportacao de arquivo `.zip` apos o processamento
- barra de progresso integrada a pagina

### Cadastros

- adicionar empresa
- editar empresa
- importar empresas por planilha `.xlsx`
- editar NSU por empresa
- resetar NSUs
- excluir empresa ou limpar todos os cadastros
- importar certificado `.pfx`

### Configuracoes

- prefixo do arquivo
- delay entre consultas
- timeout
- modo de consulta
- modo de salvamento
- opcao de baixar PDFs

## Fluxo recomendado

1. Abra `Configuracoes` e confirme os parametros do ambiente.
2. Cadastre ou importe as empresas em `Cadastros`.
3. Revise certificado, senha e vencimento das empresas.
4. Va para `Download`, selecione o periodo e as empresas desejadas.
5. Execute o download e, se necessario, exporte o `.zip` ao final.

## Arquivos principais

- `download_nfse_qt.py`: launcher da interface atual
- `ui_qt/`: camada visual em Qt
- `downloader/`: regras de download e processamento
- `config/`: configuracao, banco e persistencia
- `docs/qt_primeira_execucao.md`: guia rapido de primeira execucao

## Observacoes

- O projeto continua dependendo de configuracao valida de certificado e acesso ao Portal Nacional.
- O executavel `onefile` usa `%AppData%\\Portal NFSe` para arquivos internos e a area de trabalho para os downloads visiveis.
- Falhas de ambiente, certificado ou dependencias podem afetar o fluxo completo.
- Em caso de erro, registre a tela, a acao executada e o traceback para facilitar o diagnostico.
