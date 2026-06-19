# Monitora SGC - Automação de Dados PNCP

## Sobre o Projeto
Um pipeline de dados automatizado desenvolvido em Python para otimizar o acompanhamento de processos licitatórios. 
O script consome a API do Big Data governamental, processa históricos complexos em JSON e atualiza automaticamente um painel de indicadores no Google Sheets.

## Principais Funcionalidades
* **Extração e Tratamento (ETL):** Desempacota dados aninhados (JSON) da API pública do Estado e transforma em DataFrames tabulares usando Pandas.
* **Motor de Regras de Negócio:** Algoritmo que interpreta atrasos sistêmicos, homologações parciais e desertos/fracassos, gerando datas precisas para métricas de tempo de resposta.
* **Memória Imutável:** Sistema de caching local via Excel que protege o histórico de indicadores contra atualizações retroativas e inconsistências do banco de dados oficial.
* **Autenticação Individual:** Cofre de token local (`meu_token.txt`) que garante que cada usuário acesse a API com sua própria assinatura digital de servidor.
* **Integração Cloud:** Sincronização automática com o Google Sheets via Google Drive API para dashboards colaborativos.

## Tecnologias Utilizadas
* Python 3.x
* Pandas (Manipulação e estruturação de dados)
* Requests & JSON (Comunicação com API REST)
* Google API Client (Upload para Google Sheets via Conta de Serviço)

## Segurança e Configuração
Para rodar este projeto, é necessário configurar as credenciais de ambiente (não inclusas neste repositório por questões de segurança):
1. O token pessoal do Big Data deve ser fornecido no primeiro acesso (o sistema gerará um arquivo `meu_token.txt` automaticamente).
2. As credenciais de serviço do Google Cloud devem ser salvas como `credenciais.json` no diretório raiz.

## Lições Aprendidas
Este projeto foi fundamental para aplicar lógicas complexas de cruzamento de dados, limpeza de caracteres inválidos (`Regex`) e modelagem de chaves únicas à prova de erros (utilizando o ID do item em vez da descrição para evitar falhas de digitação do sistema origem).