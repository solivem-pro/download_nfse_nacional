# Documentação do Baixar NFS-e Portal Nacional

## Visão Geral
- [IMPORTANTE] Funcional apenas no Windowns
- Esse é um programa que capta os arquivo [.xml] e [.pdf] NFSe do Portal Nacional a partir do NSU de cada empresa cadastrada.
- O NSU é o Número Sequencial Único, cada empresa, seja prestadora ou emissora, possui contagem própria que inicia em 1 e segue aumentando em 1 a cada nota (emitida ou tomada, mesma contagem).
- Por meio da API fornecida pelo próprio governo, o processo foi automatizado e é controlado individualmente por empresa.
- Cada empresa deve ser cadastrada manualmente e posteriormente executada definindo a competência escolhida para os arquivo serem baixados. Ao final dos processos, os arquivos podem ser exportados. Só é exportável o último período processado.
- No menu [Configurações] há algumas opções de configuração básica, como se serão baixados o arquivo [.pdf], método de busca e de exportação.
- Como se trata de um projeto em fase de teste não exime de verificação humana referente a quantidade de arquivos e dados. Qualquer fator de correção favor entrar em contato.
- É possível alterar o [vba.bas] se necessário, para replicar em massa a todas as planilhas basta executar [att_planilhas.py] que as atualizações são aplicadas a todas empresas cadastradas.
- Seguindo a mesma lógica, se for necessário alterar o relatório mãe em /packs/0 e executar [att_planilhas.py] para replicar as mudanças

-----------------

## Menus
### 1. Baixar NFSe
1. Tabela das empresas cadastradas, ordenada por ordem crescente de código, mas permite sortear pelas colunas
2. Ano e mês: escolher a competência que serão baixados os arquivos.
3. Selec. Todos: seleciona todas as empresas, para selecionar ou desmarcar individualmente basta clicar na desejada que a marcação é alternada.
4. Baixar: baixa as selecionadas. Apenas ativada quando há pelo menos uma selecionada.
5. Exportar: exporta o arquivo compacto das selecionadas. Apenas ativado quando há pelo menos uma selecionada.
6. Voltar: Volta para o menu principal.
- Ao baixar um popup de processamento é iniciado com contador do progresso.
- Ao exportar é gerado um arquivo [.zip], com:
	- As pastas com [.xml] e [.pdf]:
		- PRESTADOS
		- TOMADOS
		- EVENTOS (geralmente notas canceladas, podem ser prestados ou tomados)
	- nsu_competencia.json: mostra os registros dos NSU por competência, usar apenas caso necessário e conferência.
	- erros.txt: registra os erros durante operação de download.
	- relatório_{cod}.xlsm: relatório em planilha divida em 3 abas:
		- TOMADOS
		- PRESTADOS
		- EVENTOS

---

### 2. Cadastro
1. Adicionar: Clique em adicionar para criar uma nova empresa para controle. Devem ser completados todos os campos para registrar:
	1. Código: numérico e único pra empresa, só pode ser alterado na criação da empresa
	2. Empresa: nome de identificação da empresa
	3. CNPJ: CNPJ da empresa, possui formatação automática no campo, portanto aceita entrada formatada. Único por empresa.
	4. Importar Certificado (.pfx): é necessário escolher um certificado válido para realizar o download dos arquivos, ele é importado como cópia para pasta interna.
	5. Senha Certificado: senha de acesso do certificado importado.
2. Editar: se precisar editar alguma informação ou atualizar o cadastro da empresa. Selecione uma empresa para habilitar a edição.
3. Editar NSU: todos os registros de NSU da empresa estarão neste submenu:
	1. Tabela ordenada por competência da mais atual para a mais antiga.
	2. Adicionar: para adicionar é necessário preencher todos os campos
	3. Excluir: deleta o registro selecionado. Só habilitado ao selecionar uma competência.
	4. Excluir Todos: deleta dos registros todas as competência da empresa atual.
	5. Voltar: volta para a janela anterior.
	- Campos de edição, servem para adicionar uma competência nova ou sobrescrever uma já existente.
4. Resetar NSUs: reseta os NSUs de todas as empresas, usar somente em caso de erro persistente em várias empresas.
5. Excluir: deleta o cadastro da empresa selecionada. Só habilitado ao selecionar uma competência.
6. Excluir Todos: deleta todas as empresas cadastradas.
7. Voltar: Voltar para a janela principal

-----------

### 3. Configurações

1. Prefixo Arquivo: como os arquivos [.xml] e [.pdf] serão iniciados
2. Delay(s): tempo entre lotes, afeta bloqueios de certificado
3. Timeout(s): quantos segundos o programa esperará ao máximo para obter resposta do servidor da API
4. Modo de Consulta: se a busca será por Emissão ou Competência. Em competência ele buscará também pela emissão a fim de evitar perdas de NFSe. Busca até 6 meses a frente do solicitado.
5. Modo de Cadastros: Altera a forma com que o arquivo [.zip] é exportado por CNPJ ou Código. Versátil para integrações de sistemas.
6. Baixar PDF: se marcado baixa os arquivos [.pdf] da DANFSe. Devido a instabilidades do servidor pode ocorrer de não baixar.


[Repositório no GitHub](https://github.com/solivem-pro/download_nfse_nacional)


