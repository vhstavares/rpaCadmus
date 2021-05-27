import time

import pandas as pd
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime as dt


class RPA:
    def __init__(self) -> None:
        self.driver = webdriver.Chrome()
        self.wait = WebDriverWait(driver=self.driver, timeout=30)

    def entraPagina(self, link: str) -> None:
        try:
            print(f"Acessando o link: {link}")
            self.driver.get(link)
            self.driver.maximize_window()
            time.sleep(3)
        except Exception as e:
            print(f"Erro ao acessar o link:{link}")
            print(str(e))

    def pegaVagaLocal(self, index: int) -> tuple:
        time.sleep(5)
        jsPath = f'document.querySelector("#pfolio > div:nth-child({str(index)}) > div").scrollIntoViewIfNeeded();'
        self.driver.execute_script(jsPath)
        time.sleep(3)

        cssvaga: str = f".item:nth-child({str(index)}) h3"
        csslocal: str = f".item:nth-child({str(index)}) .local"
        nomevaga = self.driver.find_element(By.CSS_SELECTOR, cssvaga).text
        local = self.driver.find_element(By.CSS_SELECTOR, csslocal).text
        return nomevaga, local

    def pegaDetalheVaga(self, index: int) -> str:
        time.sleep(5)
        cssBtnDetalhe: str = f".item:nth-child({str(index)}) .btn"
        self.wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, cssBtnDetalhe))
        )
        self.driver.find_element(By.CSS_SELECTOR, cssBtnDetalhe).click()
        cssDescricao: str = '//*[@id="boxVaga"]/p'
        descricao: str = self.driver.find_element(By.XPATH, cssDescricao).text
        xpathVagas = '//*[@id="navbarSupportedContent-7"]/ul/li[4]/a'
        time.sleep(3)
        self.driver.find_element(By.XPATH, xpathVagas).click()
        time.sleep(5)
        return descricao


if __name__ == "__main__":
    dfExcel = pd.DataFrame(columns=["Nome", "Local", "Descrição"])
    dfLog = pd.DataFrame(columns=["Tempo", "Log"])
    # Cria RPA
    try:
        rpa = RPA()
        rpa.entraPagina("https://cadmus.com.br/vagas-tecnologia/")
        dfLog = dfLog.append(
            {
                "Tempo": dt.now().strftime("%H:%M:%S"),
                "Log": "Inicio da tarefa",
            },
            ignore_index=True,
        )
        index = 1
        finaliza: bool = False
        try:
            # Fechar o aviso de cookie
            time.sleep(3)
            rpa.driver.execute_script(
                'document.querySelector("#pfolio > div:nth-child(1) > '
                'div").scrollIntoViewIfNeeded();'
            )
        except Exception as e:
            finaliza: bool = True
            print("Sem vagas")
            dfLog = dfLog.append(
                {"Tempo": dt.now().strftime("%H:%M:%S"), "Log": "Sem Vagas"},
                ignore_index=True,
            )

        while not finaliza:
            try:
                try:
                    rpa.driver.execute_script(
                        'document.querySelector("#barraMinimizada").remove();'
                    )
                except:
                    pass
                dfLog = dfLog.append(
                    {
                        "Tempo": dt.now().strftime("%H:%M:%S"),
                        "Log": "Coletando Nome da vaga e local.",
                    },
                    ignore_index=True,
                )
                nomeVaga, local = rpa.pegaVagaLocal(index=index)
                dfLog = dfLog.append(
                    {
                        "Tempo": dt.now().strftime("%H:%M:%S"),
                        "Log": "Coletando descrição.",
                    },
                    ignore_index=True,
                )
                descricao: str = rpa.pegaDetalheVaga(index=index)
                print(nomeVaga, descricao)
                linha = {
                    "Nome": nomeVaga,
                    "Local": local,
                    "Descrição": descricao,
                }
                dfExcel = dfExcel.append(linha, ignore_index=True)
                dfLog = dfLog.append(
                    {
                        "Tempo": dt.now().strftime("%H:%M:%S"),
                        "Log": f"Informações coletadas para a vaga de {nomeVaga}",
                    },
                    ignore_index=True,
                )
                index += 1
            except Exception as e:
                finaliza: bool = True
                print(str(e))
                dfLog = dfLog.append(
                    {
                        "Tempo": dt.now().strftime("%H:%M:%S"),
                        "Log": "Finalizado a coleta de vagas",
                    },
                    ignore_index=True,
                )

        rpa.driver.close()
        rpa.driver.quit()

    except Exception as e:
        dfLog = dfLog.append(
            {
                "Tempo": dt.now().strftime("%H:%M:%S"),
                "Log": f"Erro ao iniciar o Robô: {str(e)}",
            },
            ignore_index=True,
        )

    dfExcel.to_excel(
        f"{dt.now().strftime('%H_%M_%S')}_Report_Vagas.xlsx", index=False
    )
    dfLog.to_excel(f"{dt.now().strftime('%H_%M_%S')}_Vagas.xlsx", index=False)
