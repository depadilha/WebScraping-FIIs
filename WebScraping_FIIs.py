from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
import pandas as pd
import os
import time

# Obtendo o hórario do computador.

t_inicial = time.time()

# Definindo que o browser utilizado rode em segundo plano.

option = Options()
option.headless = True

# Definindo o browser e abrindo o site de onde serão retirados os nomes dos FIIs.

browser = webdriver.Firefox(options=option, executable_path=r'geckodriver.exe')
browser.get("https://fiis.com.br/lista-de-fundos-imobiliarios/")

# Espera implícita inserida para permitir que todos os componentes do site carreguem.

browser.implicitly_wait(20)

# Obtendo a quantidade de FIIs listados.

qtdd_fiis = browser.find_element_by_xpath('//*[@id="fiis-counter"]').text
int_fiis = int(qtdd_fiis[0:3])

# Lista que armazenará o nome dos FIIs.

nomes = []

# Através de um loop com range baseado no número de FIIs listados, obtêm-se os nomes dos FIIs listados no site.

for i in range(50):
    nome = browser.find_element_by_xpath(f'//*[@id="items-wrapper"]/div[{i+1}]/a/span[1]').text
    nomes.append(nome)

# Verificação.

print(nomes)

# Listas que armazenarão todos os dados dos FIIs.

precos, dy12s, dy6s, dy3s, dy1s, pvps, liqs, segs, mands, nomesok = [], [], [], [], [], [],  [], [], [], []

"""
Obtendo todos os dados dos FIIs:
  -Primeiro abre-se o site, concatenando o nome obtido anteriormente com a url base do site.
  -Em sequência há uma verificação, que age caso o site daquele FII em específico não esteja no ar ignorando-o,
   retirando o nome do FII da lista e seguindo para a próxima iteração.
  -Os dados são coletados a partir do XPATH de cada informação.
  -Os dados são armazenados em suas respectivas listas.
"""

for i, fii in enumerate(nomes):
    if i == 50:
        break
    dados_fii = []
    browser.get(f"https://www.fundsexplorer.com.br/funds/{fii}")
    try:
        element = WebDriverWait(browser, 20).until(ec.presence_of_element_located(
            (By.XPATH, '//*[@id="funds-show"]//*[@id="stock-price"]//*[@class="price"]')))
    except TimeoutException as ex:
        continue
    preco = browser.find_element_by_xpath('//*[@id="funds-show"]//*[@id="stock-price"]//*[@class="price"]').text
    dot = browser.find_element_by_xpath('//*[@id="main-indicators-carousel"]//*[@class="flickity-page-dots"]/li[2]')
    dot.click()
    pvp = browser.find_element_by_xpath('//*[@id="main-indicators-carousel"]//*[@class="flickity-slider"]/div[7]'
                                        '//*[@class="indicator-value"]').text
    liq = browser.find_element_by_xpath('//*[@id="main-indicators-carousel"]//*[@class="flickity-slider"]/div[1]'
                                        '//*[@class="indicator-value"]').text
    seg = browser.find_element_by_xpath('//*[@id="basic-infos"]//*[@class="section-body"]/div/div[2]/ul/li[4]'
                                        '//*[@class="text-wrapper"]//*[@class="description"]').text
    mand = browser.find_element_by_xpath('//*[@id="basic-infos"]//*[@class="section-body"]/div/div[2]/ul/li[3]'
                                         '//*[@class="text-wrapper"]//*[@class="description"]').text
    try:
        element2 = WebDriverWait(browser, 20).until(ec.presence_of_element_located(
            (By.XPATH, '//*[@id="dividends"]//*[@class="table"]/tbody/tr[2]/td[5]')))
        dy12 = browser.find_element_by_xpath('//*[@id="dividends"]//*[@class="table"]/tbody/tr[2]/td[5]').text
        dy6 = browser.find_element_by_xpath('//*[@id="dividends"]//*[@class="table"]/tbody/tr[2]/td[4]').text
        dy3 = browser.find_element_by_xpath('//*[@id="dividends"]//*[@class="table"]/tbody/tr[2]/td[3]').text
        dy1 = browser.find_element_by_xpath('//*[@id="dividends"]//*[@class="table"]/tbody/tr[2]/td[2]').text
    except TimeoutException as ex:
        dy12, dy6, dy3, dy1 = 0, 0, 0, 0
    precos.append(preco)
    dy12s.append(dy12)
    dy6s.append(dy6)
    dy3s.append(dy3)
    dy1s.append(dy1)
    pvps.append(pvp)
    liqs.append(liq)
    segs.append(seg)
    mands.append(mand)
    nomesok.append(fii)

# Encerrando o browser.

browser.quit()

# Gerando o Data Frame que armazenará os dados, através de um dicionário que inclui todas as listas.

todos_FIIs = {"FII": nomesok, "Preço": precos, "DY12": dy12s, "DY6": dy6s, "DY3": dy3s, "DY1": dy1s, "P/PV": pvps,
              "Liquidez (R$)": liqs, "Segmento": segs, "Mandato": mands}

FIIsDF = pd.DataFrame(data=todos_FIIs, index=None)

# Visualização

print(FIIsDF)

# Gerando o arquivo de Excel para visualização dos dados.

with pd.ExcelWriter(os.path.join(r"C:/Users/André/Desktop", "FIIs.xlsx")) as writer:
    FIIsDF.to_excel(writer, sheet_name="Dados")

# Tempo de Execução do programa.

print("O programa demorou %s minutos para rodar" % int((time.time() - t_inicial)/60))
