import time
import asyncio
from playwright.async_api import async_playwright
import pandas as pd
from dotenv import load_dotenv
import os
from unidecode import unidecode
load_dotenv()

t1 = time.time()
planilha = pd.read_excel("base_de_dados/INATIVAR.xlsx",
                         "INATIVAR", usecols=['Cd.Produto', 'Produto'], dtype='string')


async def run(playwright):
    chromium = playwright.chromium
    browser = await chromium.launch(timeout=10000)
    context = await browser.new_context()
    page = await context.new_page()
    await page.goto(os.environ['PAGE'])
    async with context.expect_event("page") as event_info:
        await page.locator(
            "//img[@src='assets_sistemas\img\sistemas\mxm.png']").click()
    mxm = await event_info.value
    await mxm.locator("#txfUsuario").fill(os.environ['USER_NAME'])
    await mxm.locator("#txfSenha").fill(os.environ['PASSWORD'])
    await mxm.locator("#ext-gen18").click()
    time.sleep(2)
    try:
        await mxm.locator("//*[@id='conpass-tag']/div/div/div[2]/div[1]/div[1]").click()
    except:
        print("Sem pesquisa")
    await mxm.locator("#tgfBusca").fill(os.environ['PROCESSO'])
    await mxm.locator("#ext-gen33").click()
    frameSearch = mxm.frame_locator("//iframe[contains(@id,'_BUSCA__IFrame')]")
    await frameSearch.locator("//a", has_text="Produto").click()
    frameProduct = mxm.frame_locator("//iframe[contains(@id,'1022_IFrame')]")
    for index, row in planilha.iterrows():
        codigo = row['Cd.Produto']
        print(codigo)
        produtos = row['Produto']
        print(unidecode(produtos).upper())
        print(index)
        await frameProduct.locator("#hpfCodigo").fill('')
        time.sleep(1.5)
        await frameProduct.locator("#hpfCodigo").fill(str(codigo))
        time.sleep(2)
        await mxm.keyboard.press('Tab')
        time.sleep(1.5)
        await frameProduct.locator("#chkControleLiberadoParaMovimentacao").uncheck()
        inativo = "inativo - "
        description = inativo + produtos
        time.sleep(1.5)
        await frameProduct.locator("#txfDescricao").fill(
            unidecode(description).upper())
        await frameProduct.locator("#hpcCodigoGrupoCotacao").fill('')
        await mxm.keyboard.press('Tab')
        await mxm.keyboard.press('Tab')
        await mxm.keyboard.press('Tab')
        await mxm.keyboard.press('Tab')
        time.sleep(1.5)
        await frameProduct.locator("#ext-gen26").click()
        time.sleep(3)
        tempos = (time.time() - t1)/60
        if(index == 500):
            print(index, " items alterados em --->> ", tempos)
        if(index == 1000):
            print(index, " items alterados em --->> ", tempos)
        if(index == 2000):
            print(index, " items alterados em --->> ", tempos)
        if(index == 4000):
            print(index, " items alterados em --->> ", tempos)
    time.sleep(5)
    tempoExec = time.time() - t1
    print("Quantidade de alterações "+index +
          "\n tempo de execução ---->>> ", tempoExec)
    await browser.close()


async def main():
    async with async_playwright() as playwright:
        await run(playwright)
asyncio.run(main())
