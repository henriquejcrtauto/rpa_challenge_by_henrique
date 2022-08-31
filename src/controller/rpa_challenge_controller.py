import logging
import traceback
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
import pyautogui as p
import pandas as pd
from src.model.get_file import get_last_file
# from src.model.clickApi import  click2 as c


# Inicializando o arquivo de log
with open(r"C:\RPA_Testes\rpa_challenge\log.txt",'w') as f:
    pass

# - Configurando o Logging
logging.basicConfig(filename=r'C:\RPA_Testes\rpa_challenge\log.txt',level=logging.INFO, format=' %(asctime)s - %(levelname)s - %(message)s',
                    datefmt='%d/%m/%Y %I:%M:%S %p')
logging.info('Inicio do Programa')


print('INICIANDO DESAFIO RPA CHALLENGE...')
logging.info('INICIANDO DESAFIO RPA CHALLENGE...')

try:

    print('Abrindo site do RPA Challenge...')
    logging.info('Abrindo site do RPA Challenge...')

    # Atualizando o Chrome Driver de Forma Dinâmica
    driver = webdriver.Chrome(ChromeDriverManager().install()) 
    navegador = driver
    navegador.get('https://www.rpachallenge.com/')
    navegador.maximize_window()
    p.sleep(2)    

    # # Download da Planilha do Desafio
    # navegador.find_element_by_xpath('/html/body/app-root/div[2]/app-rpa1/div/div[1]/div[6]/a').click()
    # p.sleep(2)

    # obtendo o nome do último arquivo baixado
    last_file = get_last_file()

    df = pd.read_excel(last_file)
    print(df)


    list_first_name = [str(x) for x in df['First Name']]
    list_last_name = [str(x) for x in df['Last Name ']]
    list_company_name = [str(x) for x in df['Company Name']]
    list_role_in_company = [str(x) for x in df['Role in Company']]
    list_address = [str(x) for x in df['Address']]
    list_email = [str(x) for x in df['Email']]
    list_phone_number = [str(x) for x in df['Phone Number']]

    # Dando start no desafio
    navegador.find_element_by_xpath('/html/body/app-root/div[2]/app-rpa1/div/div[1]/div[6]/button').click()
    p.scroll(500)
    p.sleep(1)

    print('Iniciando o desafio...')
    logging.info('Iniciando o desafio...')
    for i in range(len(list_first_name)):

        logging.info(f'RODADA {i}...')        
        
        # Preenchendo o campo First Name
        first_name_label = p.locateCenterOnScreen(r'C:\RPA_Testes\rpa_challenge\images\first_name.png', confidence=0.75)
        if first_name_label == None:
            p.scroll(-500) # Scroll Down na página para encontrar a label
            p.sleep(0.5)
            first_name_label = p.locateCenterOnScreen(r'C:\RPA_Testes\rpa_challenge\images\first_name.png', confidence=0.75)
            if first_name_label != None:
                logging.info('Encontrou a label - First Name...')
                p.click(first_name_label.x, first_name_label.y+40) # Clicando 40 pixels abaixo da imagem da label (input filed)
                p.write(list_first_name[i])  
                p.scroll(500)    
                p.sleep(0.5)    
            else:
                logging.info('Não encontrou a label - First Name...')
                      
        else:
            logging.info('Encontrou a label - First Name...')            
            p.click(first_name_label.x, first_name_label.y+40) # Clicando 40 pixels abaixo da imagem da label (input filed)
            p.write(list_first_name[i])

        # Preenchendo o campo Last Name
        last_name_label = p.locateCenterOnScreen(r'C:\RPA_Testes\rpa_challenge\images\last_name.png', confidence=0.75)
        if last_name_label == None:
            p.scroll(-500) # Scroll Down na página para encontrar a label
            p.sleep(0.5)
            last_name_label = p.locateCenterOnScreen(r'C:\RPA_Testes\rpa_challenge\images\last_name.png', confidence=0.75)
            if last_name_label != None:
                logging.info('Encontrou a label - Last Name...')
                p.click(last_name_label.x, last_name_label.y+40) # Clicando 40 pixels abaixo da imagem da label (input filed)
                p.write(list_last_name[i])         
                p.scroll(500)     
                p.sleep(0.5)  
            else:
                logging.info('Não encontrou a label - Last Name...')
        else:
            logging.info('Encontrou a label - Last Name...')
            p.click(last_name_label.x, last_name_label.y+40) # Clicando 40 pixels abaixo da imagem da label (input filed)
            p.write(list_last_name[i])
        
        # Preenchendo o campo Company Name
        company_name_label = p.locateCenterOnScreen(r'C:\RPA_Testes\rpa_challenge\images\company_name.png', confidence=0.75)
        if company_name_label == None:
            p.scroll(-500) # Scroll Down na página para encontrar a label
            p.sleep(0.5)
            company_name_label = p.locateCenterOnScreen(r'C:\RPA_Testes\rpa_challenge\images\company_name.png', confidence=0.75)
            if company_name_label != None:
                logging.info('Encontrou a label - Company Name...')
                p.click(company_name_label.x, company_name_label.y+40) # Clicando 40 pixels abaixo da imagem da label (input filed)
                p.write(list_company_name[i])    
                p.scroll(500)   
                p.sleep(0.5)        
            else:
                logging.info('Não encontrou a label - Company Name...') 
        else:
            logging.info('Encontrou a label - Company Name...')
            p.click(company_name_label.x, company_name_label.y+40) # Clicando 40 pixels abaixo da imagem da label (input filed)
            p.write(list_company_name[i])
        
        # Preenchendo o campo Role in Company
        role_in_company_label = p.locateCenterOnScreen(r'C:\RPA_Testes\rpa_challenge\images\role_in_company.png', confidence=0.75)
        if role_in_company_label == None:
            p.scroll(-500) # Scroll Down na página para encontrar a label
            p.sleep(0.5)
            role_in_company_label = p.locateCenterOnScreen(r'C:\RPA_Testes\rpa_challenge\images\role_in_company.png', confidence=0.75)
            if role_in_company_label != None:
                logging.info('Encontrou a label - Role in Company...')
                p.click(role_in_company_label.x, role_in_company_label.y+40) # Clicando 40 pixels abaixo da imagem da label (input filed)
                p.write(list_role_in_company[i])    
                p.scroll(500)   
                p.sleep(0.5)
            else:
                logging.info('Não encontrou a label - Role in Company...')
        else:
            logging.info('Encontrou a label - Role in Company...')
            p.click(role_in_company_label.x, role_in_company_label.y+40) # Clicando 40 pixels abaixo da imagem da label (input filed)
            p.write(list_role_in_company[i])
        
        # Preenchendo o campo Address
        address_label = p.locateCenterOnScreen(r'C:\RPA_Testes\rpa_challenge\images\address.png', confidence=0.75)
        if address_label == None:
            p.scroll(-500) # Scroll Down na página para encontrar a label
            p.sleep(0.5)
            address_label = p.locateCenterOnScreen(r'C:\RPA_Testes\rpa_challenge\images\address.png', confidence=0.75)
            if address_label != None:
                logging.info('Encontrou a label - Address...')
                p.click(address_label.x, address_label.y+40) # Clicando 40 pixels abaixo da imagem da label (input filed)
                p.write(list_address[i])     
                p.scroll(500)       
                p.sleep(0.5)
            else:
                logging.info('Não encontrou a label - Address...')  
        else:
            logging.info('Encontrou a label - Address...')
            p.click(address_label.x, address_label.y+40) # Clicando 40 pixels abaixo da imagem da label (input filed)
            p.write(list_address[i])    

        # Preenchendo o campo Email
        email_label = p.locateCenterOnScreen(r'C:\RPA_Testes\rpa_challenge\images\email.png', confidence=0.75)
        if email_label == None:
            p.scroll(-500) # Scroll Down na página para encontrar a label
            p.sleep(0.5)
            email_label = p.locateCenterOnScreen(r'C:\RPA_Testes\rpa_challenge\images\email.png', confidence=0.75)
            if email_label != None:
                logging.info('Encontrou a label - Email...')
                p.click(email_label.x, email_label.y+40) # Clicando 40 pixels abaixo da imagem da label (input filed)
                p.write(list_email[i])     
                p.scroll(500)     
                p.sleep(0.5) 
            else:
                logging.info('Não encontrou a label - Email...')   
        else:
            logging.info('Encontrou a label - Email...')
            p.click(email_label.x, email_label.y+40) # Clicando 40 pixels abaixo da imagem da label (input filed)
            p.write(list_email[i])
        
        # Preenchendo o campo Phone Number
        phone_number_label = p.locateCenterOnScreen(r'C:\RPA_Testes\rpa_challenge\images\phone_number.png', confidence=0.75)
        if phone_number_label == None:
            p.scroll(-500) # Scroll Down na página para encontrar a label
            p.sleep(0.5)
            phone_number_label = p.locateCenterOnScreen(r'C:\RPA_Testes\rpa_challenge\images\phone_number.png', confidence=0.75)
            if phone_number_label != None:
                logging.info('Encontrou a label - Phone Number...')
                p.click(phone_number_label.x, phone_number_label.y+40) # Clicando 40 pixels abaixo da imagem da label (input filed)
                p.write(list_phone_number[i])       
                p.scroll(500)       
                p.sleep(0.5)
            else:
                logging.info('Não encontrou a label - Phone Number...')
        else:
            logging.info('Encontrou a label - Phone Number...')
            p.click(phone_number_label.x, phone_number_label.y+40) # Clicando 40 pixels abaixo da imagem da label (input filed)
            p.write(list_phone_number[i])
        
        # Clicando no botão submit
        navegador.find_element_by_xpath('/html/body/app-root/div[2]/app-rpa1/div/div[2]/form/input').click()
        logging.info(f'FIM DA RODADA {i}...')
        p.scroll(500)
        p.sleep(1)

                                
except Exception:

    msg_error = traceback.format_exc()
    logging.info(f'{msg_error}')
    print(f'{msg_error}')


finally:

    p.alert('Desafio Concluído!')
    print('Concluído.')
    logging.info('Concluído.')
    navegador.close()
    navegador.quit()
    