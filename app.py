import openpyxl
import os
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException, NoAlertPresentException

# Carregar a planilha Excel
workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']

# Inicializa o WebDriver com o caminho específico
driver_path = r'C:\Users\Gabriella Freitas\AppData\Local\Programs\Python\Python312\chromedriver.exe'

# Verifique se o chromedriver existe no caminho especificado
if not os.path.exists(driver_path):
    print(f"O caminho especificado para o chromedriver não existe: {driver_path}")
    exit(1)

service = Service(driver_path)

# Inicialize o WebDriver do Chrome usando o serviço
try:
    driver = webdriver.Chrome(service=service)
    print("WebDriver do Chrome iniciado com sucesso.")
except WebDriverException as e:
    print(f"Erro ao iniciar o WebDriver: {e}")
    exit(1)

# URL da página de cadastro
url_cadastro = "https://cadastro-produtos-devaprender.netlify.app/index.html"

# Abre o site
driver.get(url_cadastro)

wait = WebDriverWait(driver, 15)

# Função para colar o texto usando Selenium
def paste_text(field, text):
    field.clear()
    field.send_keys(text)

# Função para clicar em um elemento
def click_element(by, value):
    element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((by, value)))
    element.click()

# Função para verificar se a janela ainda está ativa
def is_window_active(driver):
    try:
        driver.current_window_handle
        return True
    except WebDriverException:
        return False


# Loop através das linhas da planilha
for linha in sheet_produtos.iter_rows(min_row=2):
    if not is_window_active(driver):
        print("A janela do navegador foi fechada. Encerrando o script.")
        break

    try:
        # Verifica e navega para a URL de cadastro inicial
        if driver.current_url != url_cadastro:
            driver.get(url_cadastro)
            time.sleep(1)  # Pausa de 1 segundo para garantir que a página carregue completamente

        # Nome do Produto
        nome_produto = linha[0].value
        print(f"Preenchendo nome do produto: {nome_produto}")
        campo_nome_produto = wait.until(EC.presence_of_element_located((By.ID, 'product_name')))
        paste_text(campo_nome_produto, nome_produto)
        time.sleep(1)  # Pausa de 1 segundo

        # Descrição
        descricao = linha[1].value
        print(f"Preenchendo descrição do produto: {descricao}")
        campo_descricao = driver.find_element(By.ID, 'description')
        paste_text(campo_descricao, descricao)
        time.sleep(1)  # Pausa de 1 segundo

        # Categoria
        categoria = linha[2].value
        print(f"Preenchendo categoria do produto: {categoria}")
        campo_categoria = driver.find_element(By.ID, 'category')
        paste_text(campo_categoria, categoria)
        time.sleep(1)  # Pausa de 1 segundo

        # Codigo Produto
        codigo_produto = linha[3].value
        print(f"Preenchendo código do produto: {codigo_produto}")
        campo_codigo_produto = driver.find_element(By.ID, 'product_code')
        paste_text(campo_codigo_produto, codigo_produto)
        time.sleep(1)  # Pausa de 1 segundo

        # Peso
        peso = linha[4].value
        print(f"Preenchendo peso do produto: {peso}")
        campo_peso = driver.find_element(By.ID, 'weight')
        paste_text(campo_peso, peso)
        time.sleep(1)  # Pausa de 1 segundo

        # Dimensões
        dimensoes = linha[5].value
        print(f"Preenchendo dimensões do produto: {dimensoes}")
        campo_dimensoes = driver.find_element(By.ID, 'dimensions')
        paste_text(campo_dimensoes, dimensoes)
        time.sleep(1)  # Pausa de 1 segundo

        # Botão próximo
        print("Clicando no botão 'Próximo'")
        botao_proximo = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[text()="Próximo"]')))
        botao_proximo.click()
        time.sleep(1)  # Pausa de 1 segundo
        print(f"Cadastro do produto '{nome_produto}' na primeira página concluído com sucesso.")

        # Preço
        preco = linha[6].value
        print(f"Preenchendo preço do produto: {preco}")
        campo_preco = driver.find_element(By.ID, 'price')
        paste_text(campo_preco, preco)
        time.sleep(1)  # Pausa de 1 segundo

        # Quantidade em estoque
        quantidade_em_estoque = linha[7].value
        print(f"Preenchendo quantidade em estoque do produto: {quantidade_em_estoque}")
        campo_quantidade_em_estoque = driver.find_element(By.ID, 'stock')
        paste_text(campo_quantidade_em_estoque, quantidade_em_estoque)
        time.sleep(1)  # Pausa de 1 segundo

        # Data de validade
        data_de_validade = linha[8].value
        print(f"Preenchendo data de validade do produto: {data_de_validade}")
        campo_data_de_validade = driver.find_element(By.ID, 'expiry_date')
        paste_text(campo_data_de_validade, data_de_validade)
        time.sleep(1)  # Pausa de 1 segundo

        # Cor
        cor = linha[9].value
        print(f"Preenchendo cor do produto: {cor}")
        campo_cor = driver.find_element(By.ID, 'color')
        paste_text(campo_cor, cor)
        time.sleep(1)  # Pausa de 1 segundo

        # Tamanho
        tamanho = linha[10].value
        print(f"Selecionando tamanho do produto: {tamanho}")
        campo_tamanho = driver.find_element(By.ID, 'size')
        campo_tamanho.click()
        time.sleep(1)  # Pausa de 1 segundo
        if tamanho == 'Pequeno':
            click_element(By.XPATH, '//option[text()="Pequeno"]')
        elif tamanho == 'Médio':
            click_element(By.XPATH, '//option[text()="Médio"]')
        else:
            click_element(By.XPATH, '//option[text()="Grande"]')
        time.sleep(1)  # Pausa de 1 segundo

        # Material
        material = linha[11].value
        print(f"Preenchendo material do produto: {material}")
        campo_material = driver.find_element(By.ID, 'material')
        paste_text(campo_material, material)
        time.sleep(1)  # Pausa de 1 segundo

        # Botão próximo
        print("Clicando no botão 'Próximo'")
        botao_proximo = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[text()="Próximo"]')))
        botao_proximo.click()
        time.sleep(1)  # Pausa de 1 segundo
        print(f"Cadastro do produto '{nome_produto}' na segunda página concluído com sucesso.")

        # Fabricante
        fabricante = linha[12].value
        print(f"Preenchendo fabricante do produto: {fabricante}")
        campo_fabricante = driver.find_element(By.ID, 'manufacturer')
        paste_text(campo_fabricante, fabricante)
        time.sleep(1)  # Pausa de 1 segundo

        # País de origem
        pais_origem = linha[13].value
        print(f"Preenchendo país de origem do produto: {pais_origem}")
        campo_pais_origem = driver.find_element(By.ID, 'country')
        paste_text(campo_pais_origem, pais_origem)
        time.sleep(1)  # Pausa de 1 segundo

        # Observações
        observacoes = linha[14].value
        print(f"Preenchendo observações do produto: {observacoes}")
        campo_observacoes = driver.find_element(By.ID, 'remarks')
        paste_text(campo_observacoes, observacoes)
        time.sleep(1)  # Pausa de 1 segundo

        # Código de barras
        codigo_de_barras = linha[15].value
        print(f"Preenchendo código de barras do produto: {codigo_de_barras}")
        campo_codigo_de_barras = driver.find_element(By.ID, 'barcode')
        paste_text(campo_codigo_de_barras, codigo_de_barras)
        time.sleep(1)  # Pausa de 1 segundo

        # Localização no armazém
        localizacao_armazem = linha[16].value
        print(f"Preenchendo localização no armazém do produto: {localizacao_armazem}")
        campo_localizacao_armazem = driver.find_element(By.ID, 'warehouse_location')
        paste_text(campo_localizacao_armazem, localizacao_armazem)
        time.sleep(1)  # Pausa de 1 segundo

        # Botão concluir
        print("Clicando no botão 'Concluir'")
        botao_concluir = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[text()="Concluir"]')))
        botao_concluir.click()
        time.sleep(1)  # Pausa de 1 segundo

       
        # Botão adicionar mais um
       
        botao_adicionar_mais_um = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[text()="Adicionar mais um"]')))
        botao_adicionar_mais_um.click()
        time.sleep(1)  # Pausa de 1 segundo
        print("Clicando no botão 'Adicionar mais um'")
        print(f"Produto '{nome_produto}' cadastrado com sucesso.")

    except Exception as e:
        print(f"Erro ao cadastrar o produto '{nome_produto}': {e}")

# Fecha o navegador
driver.quit()
