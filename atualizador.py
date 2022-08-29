from email.utils import format_datetime
import pyautogui
import time
from datetime import datetime
from playwright.sync_api import sync_playwright
import pandas as pd
from bs4 import BeautifulSoup
import re
import openpyxl

t1 = time.time()
planilha = pd.read_excel("base_de_dados/Produtos para atualizar.xlsx", "Produtos", usecols=['Código', 'Produtos'], dtype='string')
with sync_playwright() as p:
      browser = p.chromium.launch(headless=False, timeout=5000)
      context = browser.new_context()
      page = context.new_page()
      page.goto("https://sistemas.mt.senac.br/")
      with context.expect_event("page") as event_info:
          page.locator("//img[@src='assets_sistemas\img\sistemas\mxm-treinamento.png']").click()
      mxm = event_info.value
      mxm.locator("#txfUsuario").fill("WLIK.SILVA")
      mxm.locator("#txfSenha").fill("Alterar@2022")
      mxm.locator("#ext-gen19").click()
      avaliacao = pyautogui.locateOnScreen('Avaliacao.png', confidence = 0.9)
      if (avaliacao != None):
        mxm.locator("//*[@id='conpass-tag']/div/div/div[2]/div[1]/div[1]").click()
      time.sleep(1)
      mxm.locator("#tgfBusca").fill("1022")
      mxm.locator("#ext-gen37").click()
      time.sleep(2)
      produto = pyautogui.locateOnScreen('produto.png', confidence = 0.8)
      botao = pyautogui.moveTo(produto)
      time.sleep(1)
      pyautogui.click()
      time.sleep(3)
      pesquisar = pyautogui.locateOnScreen('pesquisar.png', confidence = 0.8)
      time.sleep(2)
      pyautogui.moveTo(x= pesquisar.left+85, y= pesquisar.top+4)
      time.sleep(1)
      mxm.locator("#ext-gen877").click()
      product = mxm.frame("1022_IFrame")
      for index,row in planilha.iterrows():
        codigo = row['Código']
        print(str(codigo))
        produtos = row['Produtos']
        print(str(produtos))
        print(index)
        product.fill("#hpfCodigo", "")
        product.fill("#hpfCodigo", str(codigo))
        # time.sleep(1)
        mxm.keyboard.press('Tab')
        # time.sleep(1)
        product.locator("#chkControleLiberadoParaMovimentacao").uncheck()
        inativo = "inativo - "
        product.locator("#txfDescricao").fill(inativo.upper()+str(produtos).upper())
        # time.sleep(1)
        tipo_do_item = "07 - MATERIAL DE USO E CONSUMO"
        product.locator("#hpcTipoItem").fill(tipo_do_item)
        # time.sleep(0.5)
        mxm.keyboard.press('Tab')
        # time.sleep(1)
        product.locator("#ext-gen26").click()
        time.sleep(3)
      time.sleep(5)
      tempoExec = time.time() -t1
      print("Quantidade de alterações "+index+"\n tempo de execução ---->>> ", tempoExec)
      browser.close()