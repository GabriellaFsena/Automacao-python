## Automação de Cadastro de Produtos em Formulário Web
Este projeto consiste em um script Python para automação do cadastro de produtos em um formulário web. Ele utiliza a biblioteca Selenium para interagir com o navegador Chrome e Openpyxl para ler dados de uma planilha Excel.

## Pré-requisitos
Python 3.x instalado
Bibliotecas Python: openpyxl, selenium
Certifique-se de ter o Chrome WebDriver configurado corretamente no seu sistema. O WebDriver pode ser baixado aqui.

## Instalação das Dependências
Instale as bibliotecas necessárias usando o pip:

bash
Copiar código
pip install openpyxl selenium
## Como Usar
Clone este repositório:
bash
Copiar código
git clone https://github.com/GabriellaFsena/Automacao.git
cd nome-do-repositorio
Coloque sua planilha Excel (produtos_ficticios.xlsx) na raiz do projeto. Certifique-se de que ela tenha as seguintes colunas: Nome do Produto, Descrição, Categoria, Código do Produto, Peso, Dimensões, Preço, Quantidade em Estoque, Data de Validade, Cor, Tamanho, Material, Fabricante, País de Origem, Observações, Código de Barras, Localização no Armazém.

Execute o script Python cadastro_produtos.py:

bash
Copiar código
python cadastro_produtos.py
O script começará a preencher o formulário web com os dados de cada produto na planilha, seguindo os seguintes passos:

Preenchimento dos campos do produto na primeira página do formulário.
Clique no botão "Próximo".
Preenchimento dos campos adicionais do produto na segunda página.
Clique no botão "Próximo".
Preenchimento dos campos finais do produto na terceira página.
Clique no botão "Concluir".
Aceitação de alertas de confirmação.
Clique no botão "Adicionar Mais Um" para cada novo produto na planilha.
Notas
A automação pode falhar se houver problemas de conexão com o navegador ou se o formato da planilha estiver incorreto.
Certifique-se de que a URL do formulário web está correta e acessível durante a execução do script.
Ajuste os tempos de espera (time.sleep()) conforme necessário para garantir que o script funcione corretamente em seu ambiente.
