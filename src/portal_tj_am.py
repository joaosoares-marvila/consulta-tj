from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from loguru import logger
from selenium.webdriver.chrome.options import Options
from time import sleep
import traceback
from planilha import Planilha
import os
from datetime import datetime
import pyautogui


diretorio_atual = os.getcwd()
caminho_planilha_processos = os.path.join(diretorio_atual, 'assets', 'processos_tja.xlsx')



class PortalTJAM:

    def __init__(self, driver: webdriver.Chrome) -> None:
        self.driver = driver
     
        cabecalho_processos = ['processo', 'url', 'requerente', 'situacao', 'classe', 'assunto', 'foro', 'vara', 'juiz', 'data_ultima_movimentacao', 'ultima_movimentacao', 'valor', 'horario']
        self.planilha_processos = Planilha(
            caminho_arquivo=caminho_planilha_processos,
            nome_planilha='TJAM - PROCESSOS',
            cabecalho=cabecalho_processos
        )



    def principal(self) -> None:

        url_pesquisa = 'https://consultasaj.tjam.jus.br/cpopg/search.do;jsessionid=929486DF9959C512A18456F60BBAF9CF.cpopg1?conversationId=&cbPesquisa=DOCPARTE&dadosConsulta.valorConsulta=11669325000188&cdForo=-1'
        urls_processos = self.obtem_processos(url_pesquisa=url_pesquisa)

        
        for url_processo in urls_processos:
        
            processos_ja_obtidos_planilha = self.planilha_processos.obtem_url_processos()

            processo_obtido = {}
            for processo in processos_ja_obtidos_planilha:
                if processo['url'] == url_processo:
                    processo_obtido = processo
                    break
                    
            if processo_obtido: 
                logger.warning(f"Processo {processo['processo']} já obtido")
                self.planilha_processos.define_horario_verificacao_processo(numero_processo=processo['processo'])
                continue
                    
            dados_processo = self.__obter_dados_processo(url_processo=url_processo)
            dados_processo['horario'] = datetime.now().strftime("%m/%d/%Y, %H:%M:%S")
                    
            self.planilha_processos.adiciona_dado(dados=dados_processo)
            
            


    def obtem_processos(self, url_pesquisa: str) -> list:

        sleep(2)

        logger.success(f"Obtendo processo do TJAM")
        self.driver.get(url=url_pesquisa)

        # AGUARDA URL CARREGAR
        while self.driver.current_url != url_pesquisa:
            logger.warning("Aguardando página carregar")
            sleep(1)
        
        if 'processo.numero' not in url_pesquisa:
          
            # Número de processos totais no site referente ao CNPJ da ympactus
            processos_totais = WebDriverWait(self.driver, 10).until(
                        EC.visibility_of_element_located((By.XPATH, "(//span[contains(@id, 'contadorDeProcessos')])[1]"))).text.strip().split(' ')[0]
            processos_totais = int(processos_totais)


            # Número de processos que já foram obtidos
            processos_obtidos = 0

            links_processos = []

            while processos_obtidos != processos_totais:

                # Cards dos proessos
                processos = WebDriverWait(self.driver, 10).until(
                        EC.visibility_of_all_elements_located((By.XPATH, "//div[contains(@id, 'listagemDeProcessos')]//li//a[contains(@class,'linkProcesso')]")))
                
                # Itera sobre os processos
                for processo in processos:

                    link_processo = processo.get_attribute('href')
                    links_processos.append(link_processo)
                    processos_obtidos += 1


                # Avança para a próxima página
                if processos_obtidos != processos_totais:
            
                    btn_avancar_pagina = WebDriverWait(self.driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "(//a[contains(@class, 'unj-pagination__next icon-next')])[1]")))
                    btn_avancar_pagina.click()
                    sleep(5)

            return links_processos

        # PROCESSO UNICO
        else:
            return [url_pesquisa]




    def __obter_dados_processo(self, url_processo:str) -> dict:

        self.driver.get(url_processo)
        logger.info("---------------------------------")
        
        dados = {}
        
        # URL
        dados['url'] = url_processo

        # processo
        processo = WebDriverWait(self.driver, 100).until(
            EC.visibility_of_element_located((By.XPATH, f"//span[contains(@id,'numeroProcesso')]"))).text.strip()
        dados['processo'] = processo
        logger.info(f"Processo: {processo}")
        
        # Autor / demandante / requerente
        try:
            requerente = WebDriverWait(self.driver, 3).until(
                    EC.visibility_of_element_located((By.XPATH, "(//span[contains(text(),'Autor') or contains(text(),'Exequente') or contains(text(),'Credor') or contains(text(),'Liquidante')or contains(text(),'Reqte') or contains(text(), 'Requerente') or contains(text(), 'Exeqte') or contains(text(),'Demandante') ])[1]/parent::*//..//td[2]"))).text.strip().split('\n')[0]
            dados['requerente'] = requerente
            logger.info(f"Requerente: {requerente}")
        except:
            logger.error("Sem autor")
            dados['requerente'] = '-'

        # Situação
        try:
            situacao = WebDriverWait(self.driver, 2).until(
                    EC.visibility_of_element_located((By.XPATH, "//span[contains(@id, 'labelSituacaoProcesso')]"))).text.strip()
            dados['situacao'] = situacao
            logger.info(f"Situação: {situacao} ")
        except:
            dados['situacao'] = '-'
            logger.warning("Sem situação")

        # Classe
        classe = WebDriverWait(self.driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, "//span[contains(@id, 'classeProcesso')]"))).text.strip()
        dados['classe'] = classe
        logger.info(f"Classe: {classe}")
        
        # Assunto
        assunto = WebDriverWait(self.driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, "//span[contains(@id, 'assuntoProcesso')]"))).text.strip()
        dados['assunto'] = assunto
        logger.info(f"Assunto: {assunto}")
        
        # Foro
        foro = WebDriverWait(self.driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, "//span[contains(@id, 'foroProcesso')]"))).text.strip()
        dados['foro'] = foro
        logger.info(f"Foro: {foro}")

        # Vara
        vara = WebDriverWait(self.driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, "//span[contains(@id, 'varaProcesso')]"))).text.strip()
        dados['vara'] = vara
        logger.info(f"Vara: {vara}")

        # Juiz
        try:
            juiz = WebDriverWait(self.driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, "//span[contains(@id, 'juizProcesso')]"))).text.strip()
            dados['juiz'] = juiz
            logger.info(f"Juiz: {juiz}")
        except:
            dados['juiz'] = '-'

        # Data ultima movimentação
        data_ultima_movimentacao = WebDriverWait(self.driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, "(//td[contains(@class,'dataMovimentacao')])[1]"))).text.strip()
        dados['data_ultima_movimentacao'] = data_ultima_movimentacao
        logger.info(f"Data da ultima movimentação: {data_ultima_movimentacao}")
         
        # Ultima Movimentação
        ultima_movimentacao = WebDriverWait(self.driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, "(//td[contains(@class,'descricaoMovimentacao')])[1]"))).text.strip()
        dados['ultima_movimentacao'] = ultima_movimentacao
        logger.info(f"Ultima movimentação: {ultima_movimentacao}")

        # Mais informações (valor)
        try:
            mais_detalhes = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//a[contains(@href, 'maisDetalhes')]")))
            mais_detalhes.click()

            valor = WebDriverWait(self.driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, "//div[contains(@id, 'valorAcaoProcesso')]"))).text.strip()
    
            dados['valor'] = valor
            logger.info(f"Valor: {valor}")
        
        except TimeoutException:
            logger.warning("Sem valor na ação")
            dados['valor'] = '-'
        except Exception as error:
            logger.error(f"Não foi possível obter o valor da ação | {str(error)}")
            dados['valor'] = '-'
            
            
        return dados
        


if __name__ == '__main__':
    
    options = Options()
    options.add_experimental_option('excludeSwitches', ['enable-logging', 'enable-automation'])
    # options.add_argument('--headless')
    driver = webdriver.Chrome(options=options)
    driver.set_window_size(width=1920,height=1080)


    #driver = uc.Chrome(headless=False,use_subprocess=False)

    portal_tj_am = PortalTJAM(driver=driver)

    portal_tj_am.principal()
    #portal_tj_sp.main()
    
    
    #portal_tj_ac.principal()
    
