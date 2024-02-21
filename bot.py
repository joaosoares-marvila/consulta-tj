from loguru import logger
import undetected_chromedriver as uc
from planilha import Planilhas
from src.portal_jusbrasil import PortalJusBrasil
from src.portal_tjes import PortalTJES
import traceback

def main():
    
    planilhas = Planilhas()
    driver = uc.Chrome(headless=False,use_subprocess=False)
    portal_tjes = PortalTJES(
        driver=driver
    )
    portal_tjes.abre_pag_pesquisa()
    
    while True:

        # Obtém o próximo credor a ser consultado
        try:
            cpf, credor = planilhas.obter_proximo_credor()
        
            try:
                """processo_obtido = portal_jusbrasil.obter_processo(
                    cpf_credor=cpf
                )
                 """
                
                credor = credor.replace('"','')
                dados_processo = portal_tjes.main(nome_credor=credor)
                
                if dados_processo:
                    planilhas.preencher_dados_processo(
                        nome_credor = credor,
                        dados_processo=dados_processo
                    )
                else:
                    raise Exception('Nenhum dado obtido')


                
                planilhas.preenche_log(
                    cpf=cpf,
                    retorno='Sucesso'
                )
            except Exception as error:
                logger.error(str(error))

                if 'Nenhum processo vinculado ao CPF pelos tribunais' in str(error) or 'Nenhum processo vinculado a Ympactus' in str(error):
                    dados_processo = {
                        'numero': 'SEM',
                        'valor_causa': '',
                        'natureza' : ''
                    }
                    planilhas.preencher_dados_processo(
                        nome_credor=credor,
                        dados_processo=dados_processo
                    )
                    planilhas.preenche_log(
                        cpf=cpf,
                        retorno='Sucesso'
                    )
                else:
                    planilhas.preencher_dados_credores_erros(
                        cpf=cpf,
                        nome_credor=credor,
                        erro=str(error)
                    )
                    planilhas.preenche_log(
                        cpf=cpf,
                        retorno='Erro'
                    )

        # Quebra a execução caso não seja possível obter o próximo credor
        except Exception as error:
            logger.error(traceback.format_exc())
            logger.warning("----------------------------")
            logger.error(str(error))
            logger.warning("----------------------------")
            break
                

if __name__ == '__main__':
    main()
