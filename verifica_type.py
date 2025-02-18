from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import dotenv
import os
import pandas as pd
import time


dotenv.load_dotenv()

options = webdriver.ChromeOptions()
options.add_argument("--force-device-scale-factor=0.8")

browse = webdriver.Chrome(options=options)
browse.get("https://evo5.w12app.com.br/#/acesso/bluefit/autenticacao")
browse.maximize_window()

login_user = browse.find_element("id", "usuario")
login_user.clear()
login_user.send_keys(os.environ["EMAIL"])

login_pass = browse.find_element("id", "senha")
login_pass.clear()
login_pass.send_keys(os.environ["PASSWORD"])

button_login = browse.find_element("id", "entrar")
button_login.click()

# Em spreadsheet_pay coloque o nome da lista no formato xlsx 
spreadsheet_pay = 'nome_da_lista.xlsx'
df = pd.read_excel(spreadsheet_pay, usecols=[0], skiprows=303, nrows=149, dtype={0: int})
# Utilize o parâmetro dtype={0: int} quando o número da coluna não for reconhecido como int.

# Caso possua algum id que não necessite de cobrança, coloque-o dentro da lista exclude_numbers.
exclude_numbers = [414141]
df = df.loc[~df.iloc[:, 0].isin(exclude_numbers)]

sucess = 0
no_pending = 0
fails = 0
time.sleep(3)

for index, row in df.iterrows():
    id_value = row.iloc[0]
    max_attempts = 5
    attempts = 0
    success = False

    while attempts < max_attempts and not success:
        try:
            waiting = WebDriverWait(browse, 6)
            
            id_search = browse.find_element("id", "evoAutocomplete")
            id_search.clear()
            id_search.send_keys(str(id_value))
            time.sleep(3)

            id_desired = browse.find_element("id", "cdk-overlay-0")
            waiting.until(EC.element_to_be_clickable(id_desired))
            id_desired.click()

            try:
                pending_exists = waiting.until(EC.visibility_of_element_located(
                    (By.CSS_SELECTOR, ".full.border-radius-5.m-b-lg.bg-green.text-green.border-green.ng-star-inserted"))
                )
                if "Cliente sem pendências" in pending_exists.text or "Acesso liberado" in pending_exists.text:
                    print(f"ID {id_value}: Cliente sem pendências")
                    no_pending += 1
                    success = True
                    break
            except TimeoutException:
                pass

            pending = browse.find_element(By.CSS_SELECTOR, "[data-cy='recebimento-devedor']")
            pending.click()
            time.sleep(2)

            button_select_pay = browse.find_element(By.CSS_SELECTOR,
                ".mat-focus-indicator.mat-icon-button.mat-button-base.icone-small.evo-button.tertiary.m-x-xs.ng-star-inserted")
            waiting.until(EC.element_to_be_clickable(button_select_pay))
            button_select_pay.click()
            time.sleep(2)

            button_online = browse.find_element(By.XPATH, "//md-radio-button[@role='radio' and @aria-label='On-line']")
            waiting.until(EC.element_to_be_clickable(button_online))
            button_online.click()

            select_payment = waiting.until(EC.element_to_be_clickable((By.NAME, "clienteCartao")))
            select_payment.click()

            select_last_card = browse.find_elements(By.ID, "idCartaoOnline")
            if select_last_card:
                ActionChains(browse).move_to_element(select_last_card[-1]).click().perform()

            select_vezes = waiting.until(EC.element_to_be_clickable((By.NAME, "totalParcelas")))
            select_vezes.click()
            time.sleep(1)

            select_one = waiting.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'md-option[value="1"]')))
            select_one.click()

            button_add = browse.find_element(By.ID, "adicionar")
            waiting.until(EC.element_to_be_clickable(button_add))
            button_add.click()
            time.sleep(1)

            try:
                button_window_confirm = browse.find_element(By.XPATH, '//button[@ng-click="vm.fechar(true)"]')
                waiting.until(EC.element_to_be_clickable(button_window_confirm))
                button_window_confirm.click()
            except NoSuchElementException:
                print("...")

            time.sleep(1)

            button_finalizar = browse.find_element(By.ID, "finalizar")
            waiting.until(EC.element_to_be_clickable(button_finalizar))
            button_finalizar.click()
            time.sleep(5)

            try:
                div_rejeitado = browse.find_element(By.CSS_SELECTOR, ".text-center.texto.vermelho.ng-binding")
                if "foi rejeitado pela operadora" in div_rejeitado.text or "(Cod. 80)" in div_rejeitado.text or "(Cod. 56)" in div_rejeitado.text:
                    print(f"ID {id_value}: Pagamento rejeitado")
                    fails += 1
                else:
                    sucess += 1
                    print(f"ID {id_value}: Sucesso no pagamento")
                success = True
            except NoSuchElementException:
                print(f"ID {id_value}: Algo inesperado aconteceu.")
                fails += 1

            print(f"ID Atual: {id_value}\nSuccess: {sucess} - No pending: {no_pending} - Fails: {fails}")

        except (NoSuchElementException, TimeoutException) as e:
            print(f"Tentativa {attempts + 1} falhou para ID {id_value}: {str(e)}")
            attempts += 1
            time.sleep(1)

    if not success:
        print(f"ID {id_value}: Todas as tentativas falharam.")
        fails += 1

print(f"Total de pagamentos realizados com sucesso: {sucess}")
