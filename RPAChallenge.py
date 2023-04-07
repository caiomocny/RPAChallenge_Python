# install openpyxl

def main():

    import time
    import pandas as pd
    from selenium import webdriver
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.support.wait import WebDriverWait
    from selenium.webdriver.common.by import By

    # Read Excel file
    df = pd.read_excel("input\challenge.xlsx")
    print(df.columns)

    # Nomear as colunas de forma correta (tratar possíveis erros que possam vir da planilha baixada)
    df.columns = ['First Name', 'Last Name', 'Company Name', 'Role in Company', 'Address', 'Email', 'Phone Number']

    # Transformar o tipo de objeto da variável "Phone Number" para string de modo que possa ser passado como texto no
    # site
    df["Phone Number"] = df["Phone Number"].astype(str)

    print(df.info())
    print(df)
    print(f"A tabela possui {len(df)} registros.")

    # Acessar o Chrome driver
    service_driver = Service(executable_path='chromedriver.exe')
    driver = webdriver.Chrome(service=service_driver)

    # Garantir que o Chrome esteja pronto para ser utilizado
    driver.implicitly_wait(0.5)

    # Acessar o site RPA Challenge
    driver.get("http://rpachallenge.com/")
    driver.maximize_window()

    # Wait "Start button" be clickable
    start_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//button')))
    time.sleep(2)
    start_button.click()

    # Inserir os dados nos devidos campos
    for person_row in range(len(df)):  # Para cada registro do dataframe
        for field_column in range(len(df.columns)):  # Para cada coluna do registro vigente
            # //rpa1-field -> elemento pai âncora que possui o atributo fiel @ng-reflect-dictionary-value que por sua
            # vez contém o valor da variável que estou trabalhando (coluna do dataframe) //input -> elemento filho
            # que está ancorado
            driver.find_element(By.XPATH,
                                '//rpa1-field[@ng-reflect-dictionary-value="' + df.columns[field_column] + '"]//input') \
                .send_keys(df.iloc[person_row, field_column])

        driver.find_element(By.XPATH, '//input[@class="btn uiColorButton"]').click()
        time.sleep(1)

    time.sleep(1)

    # Extrair o tempo de processamento do processo
    result = driver.find_element(By.XPATH, '//*[@class="message2"]').text
    result = result[result.find("in") + 3:result.find("milli")]

    # Manipular a variável resultado
    result = float(result)/1000
    result = "{0} segundos".format(result)

    return result
