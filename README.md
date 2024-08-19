# Geração de Documentos de Pesquisa com Google Docs

Este script do Google Apps Script automatiza a criação de documentos de pesquisa usando dados de uma planilha do Google Sheets. Ele gera um documento no Google Docs a partir de um modelo predefinido e preenche o conteúdo com base nas informações da planilha, incluindo uma tabela de pesquisa de preços.
Funcionalidade

    Leitura de Dados: Obtém dados da planilha ativa no Google Sheets, incluindo informações gerais e os dados da pesquisa.
    Criação do Documento: Cria uma cópia de um documento modelo existente no Google Drive e o renomeia conforme especificado.
    Substituição de Marcadores: Substitui marcadores no documento modelo com dados da planilha, como obra, endereço, agente, função, posto e data atual.
    Preenchimento da Tabela: Preenche uma tabela no documento com informações sobre equipamentos e preços de diferentes empresas, calculando a mediana dos preços.
    Salvar e Compartilhar: Salva e fecha o documento criado e atualiza a planilha com o link para o novo documento.

## Aba start 
![image](https://github.com/user-attachments/assets/53cbd8d3-0e57-4575-8e14-42a81ce3b97e)

## Aba pesquisa
![image](https://github.com/user-attachments/assets/7d3e450d-8296-4b73-b713-866bf0e10a9d)
# Como Usar

    Configuração:
        Abra o Google Sheets onde você deseja usar o script.
        Crie uma aba chamada "Pesquisa" e insira seus dados conforme as colunas esperadas.
        Crie uma aba chamada "Start" e insira as informações necessárias nas células B5 a B10.





    Script:
        Abra o Editor de Scripts do Google Apps (Extensões > Apps Script).
        Copie e cole o código fornecido no Editor de Scripts.
        Atualize os IDs do template e da pasta de destino conforme necessário.

    Execução:
        Execute a função createNewGoogleDocs() a partir do Editor de Scripts para gerar o documento de pesquisa.
        O link para o novo documento será atualizado na célula E4 da aba "Start".
