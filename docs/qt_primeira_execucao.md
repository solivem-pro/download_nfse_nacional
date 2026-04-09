# Primeira Execucao

## Preparacao

- Instale as dependencias do projeto com `pip install -r requirements.txt`.
- Garanta que o `python` usado para abrir a interface seja o mesmo ambiente onde as dependencias foram instaladas.
- O pos-processamento agora e nativo em Python, entao nao depende de Excel na maquina.

## Como abrir

Execute:

```powershell
python download_nfse_qt.py
```

## Smoke test sugerido

1. Abra `Configuracoes` e confirme se os valores atuais carregam de `config/config.json`.
2. Salve uma alteracao simples em `Configuracoes` e confirme se ela persiste ao reabrir a tela.
3. Abra `Cadastros`, selecione uma empresa existente e teste `Editar`.
4. Em `Cadastros`, teste `Editar NSU` e salve um registro de exemplo.
5. Em `Cadastros`, teste `Importar XLSX` com uma planilha no formato esperado.
6. Em `Download`, selecione uma empresa com certificado valido e inicie o fluxo.
7. Ao final do download, confirme se:
   - a barra de progresso acompanha a execucao
   - o resumo final abre corretamente
   - o `.zip` da empresa foi gerado em `packs/<codigo>.zip`

## Pontos externos que podem falhar

- certificado `.pfx` invalido ou senha incorreta
- `openssl` ausente ao tentar fallback de leitura de vencimento
- ambiente Python diferente do ambiente onde as dependencias foram instaladas

## Se der erro

Envie:

- o traceback completo
- a tela onde aconteceu
- o botao ou acao que disparou o erro
- se o erro ocorreu ao abrir a interface ou durante `Download` / `Cadastros`
