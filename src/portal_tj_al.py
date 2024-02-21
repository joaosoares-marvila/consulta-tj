from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from loguru import logger
from selenium.webdriver.chrome.options import Options
from time import sleep
from planilha import Planilha
import os


class PortalTJAL:

    def __init__(self, driver) -> None:
        self.driver = driver

    def __obter_processos(self)-> dict:

        # Página 1 da consulta pelo cnpj
        
        self.driver.get('https://www2.tjal.jus.br/cpopg/trocarPagina.do?paginaConsulta=1&paginaConsulta=2&paginaConsulta=4&paginaConsulta=5&conversationId=&cbPesquisa=DOCPARTE&dadosConsulta.valorConsulta=11669325000188&cdForo=-1')

        # Número de processos totais no site referente ao CNPJ da ympactus
        processos_totais = WebDriverWait(self.driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, "(//span[contains(@id, 'contadorDeProcessos')])[1]"))).text.strip().split(' ')[0]
        processos_totais = int(processos_totais)


        # Número de processos que já foram obtidos
        processos_obtidos = 0

        # Dicionário que conterá o número do processo como chave e o link como valor
        links_processos = {}

        while processos_obtidos != processos_totais:

            # Cards dos proessos
            processos = WebDriverWait(self.driver, 10).until(
                    EC.visibility_of_all_elements_located((By.XPATH, "//div[contains(@id, 'listagemDeProcessos')]//li//a[contains(@class,'linkProcesso')]")))
            
            # Itera sobre os processos
            for processo in processos:

                link_processo = processo.get_attribute('href')
                numero_processo = processo.text
                links_processos[numero_processo] = link_processo

                processos_obtidos += 1


            if processos_obtidos != processos_totais:
                # Avança para a próxima página
                btn_avancar_pagina = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "(//a[contains(@class, 'unj-pagination__next icon-next')])[1]")))
                btn_avancar_pagina.click()
                sleep(5)

        return links_processos
    
    def __obter_dados_processo(self, url_processo:str, processo: str) -> None:

        self.driver.get(url_processo)
        
        dados = {}

        WebDriverWait(self.driver, 100).until(
                    EC.visibility_of_element_located((By.XPATH, f"(//span[contains(text(),'{processo}')])[1]"))).text.strip()
        
        
         # Autor / demandante / requerente
        try:
            requerente = WebDriverWait(self.driver, 3).until(
                    EC.visibility_of_element_located((By.XPATH, "(//span[contains(text(),'Autor') or contains(text(),'Exequente') or contains(text(), 'Requerente') or contains(text(),'Demandante') ])[1]/parent::*//..//td[2]"))).text.strip().split('\n')[0]
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
        juiz = WebDriverWait(self.driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, "//span[contains(@id, 'juizProcesso')]"))).text.strip()
        dados['juiz'] = juiz
        logger.info(f"Juiz: {juiz}")

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
        
    def teste(self) -> None:
        self.__obter_dados_processo(url_processo='https://www2.tjal.jus.br/cpopg/show.do?processo.codigo=1M0001XJS0000&processo.foro=58&paginaConsulta=2&paginaConsulta=1&paginaConsulta=2&paginaConsulta=4&paginaConsulta=5&conversationId=&cbPesquisa=DOCPARTE&dadosConsulta.valorConsulta=11669325000188&cdForo=-1')


    def main(self):
        diretorio_atual = os.getcwd()
        caminho_planilha = os.path.join(diretorio_atual, 'assets', 'processos.xlsx')
        
        
        cabecalho = ['processo', 'url', 'requerente', 'situacao', 'classe', 'assunto', 'foro', 'vara', 'juiz', 'data_ultima_movimentacao', 'ultima_movimentacao', 'valor']
        planilha = Planilha(
            caminho_arquivo=caminho_planilha,
            nome_planilha='TJAL',
            cabecalho=cabecalho
        )

        processos_planilha = planilha.obtem_processos_planilha()
        
        links_processos = self.__obter_processos()

        for processo, url in links_processos.items():

            if processo in processos_planilha:
                logger.warning(f"Processo {processo} já registrado")
                continue
            
            logger.warning(f"-----------------------------------------------")
            logger.warning(f"Obtendo dados do processo {processo}")
            logger.info(f"Processo: {processo}")
            logger.info(f"URL: {url}")

            dados_processo = self.__obter_dados_processo(url_processo=url, processo=processo)
            

            dados_processo['processo'] = processo
            dados_processo['url'] = url

            planilha.adiciona_processo(dados_processo=dados_processo)



if __name__ == '__main__':
    
    options = Options()
    options.add_experimental_option('excludeSwitches', ['enable-logging', 'enable-automation'])
    # options.add_argument('--headless')
    driver = webdriver.Chrome(options=options)
    driver.set_window_size(width=1920,height=1080)


    #driver = uc.Chrome(headless=False,use_subprocess=False)

    portal_tj_al = PortalTJAL(driver=driver)

    portal_tj_al.main()
    

    
