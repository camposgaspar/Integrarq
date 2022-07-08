import selenium.webdriver as webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from datetime import datetime, timedelta
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import ElementNotInteractableException
import time


class Module_SOLUCIONARE():

    def __init__(self):
        chrome_options = Options()
        chrome_options.add_experimental_option("detach", True)
        chrome_options.add_argument("--window-size=1920,1080")
        self.default_text = "Mensagem automática. Caso precise de ajuda, por favor entre em contato conosco pelo " \
                            "e-mail antendimento@integrarq.com.br"
        self.url = "http://online.solucionarelj.com.br:9191/gerenciador/login.php"

    @staticmethod
    def init_driver():
        """Inicia e retorna o Webdriver"""
        return webdriver.Chrome()

    @staticmethod
    def close_driver(driver):
        """Encerra o Webdriver.

        Parameters:
            driver (object): Webriver que será encerrado
            """
        driver.quit()

    @staticmethod
    def login_solucionare(driver, url, company, user, password):
        """Login no sistema.

        Parameters:
            driver (object):    Selenium Webdriver
            url (str):  URL do portal de login da Solucionare
            company (str):  Empresa - Parte 1 do Login
            user (str): Usuário - Parte 2 do login
            password (str): Senha - Parte 3 do login
            """
        driver.implicitly_wait(10)
        driver.get(url)

        # Fill login form
        driver.find_element(By.NAME, "CODIGO").send_keys(company)
        driver.find_element(By.NAME, "USUARIO").send_keys(user)
        driver.find_element(By.NAME, "SENHA").send_keys(password)
        # Submit
        driver.find_element(By.NAME, "Submit").click()

    @staticmethod
    def send_codes(box, codes):
        """Insere os códigos no campo.
            (Código; TAB; Código; TAB)

        Parameters:
            box (object):   Objeto do webdriver onde os códigos serão inseridos
            codes (list):  Códigos aos quais serão enviados os andamentos
            """
        for code in codes:
            box.send_keys(code)
            box.send_keys(Keys.TAB)

    @staticmethod
    def find_date():
        """Encontra as data para preenchimento do formulário"""
        day = datetime.today()
        month = datetime.today().strftime('%m')
        year = datetime.today().strftime('%Y')

        first_of = day.replace(day=1)
        last_month = (first_of - timedelta(days=1)).strftime('%m')

        return day.strftime('%d'), month, last_month, year

    def email_send(self, driver, codes):
        """Preenche o formulário no site.

        Parameters:
            driver(object): Webdriver
            codes (list):  Códigos aos quais serão enviados os andamentos
            """
        # Acessa área de e-mails
        driver.get(
            "http://online.solucionarelj.com.br:9191/gerenciador/gerenciador_novo/modulos/andamento_processual/router"
            ".php?controller=Email")

        # Com novos andamentos
        driver.find_element(By.XPATH, '//*[@id="form_email"]/div[1]/div[1]/label[2]/div/ins').click()

        # Caixa de códigos; insere os códigos na caixa: código, TAB, código, TAB, etc.
        box = driver.find_element(By.XPATH, '//*[@id="s2id_autogen2"]')
        self.send_codes(box, codes)

        # Com novos andamentos
        driver.find_element(By.XPATH, '//*[@id="form_email"]/div[2]/div/label[2]/div/ins').click()

        # Apenas novos a partir de data
        driver.find_element(By.XPATH, '//*[@id="form_email"]/div[4]/div/label[3]/div/ins').click()

        # TODO - TESTE - Últimos 5 movimentos / Enviar todos
        # driver.find_element(By.XPATH, '//*[@id="form_email"]/div[4]/div/label[5]/div/ins').click()
        # driver.find_element(By.XPATH, '//*[@id="form_email"]/div[2]/div/label[1]/div/ins').click()

        # Define datas, hoje -1 mês; Insere datas no campo; Clica no bootstrap para inserir a data no form.
        day, month, last_month, year = self.find_date()
        driver.find_element(By.XPATH, '//*[@id="form_email"]/div[5]/div/div[1]/div/input').send_keys(
            f"{day}/{last_month}/{year}", Keys.ESCAPE)
        driver.find_element(By.XPATH,
                            '/html/body/div[2]/aside[2]/section[2]/div/div/div/div/form/div[5]/div/div[1]/div/span').click()
        table = driver.find_element(By.XPATH, '/html/body/div[6]/div[3]/table')
        driver.find_element(By.XPATH,
                            '/html/body/div[2]/aside[2]/section[2]/div/div/div/div/form/div[5]/div/div[1]/div/span').click()

        try:  # Windows
            day_today = datetime.today().strftime("%#d")
        except:  # Linux
            day_today = datetime.today().strftime("%-d")

        for row in table.find_elements(By.XPATH, './/td'):
            if row.text == day_today:
                row.click()
                break

        # E-mails do escritório
        driver.find_element(By.XPATH, '//*[@id="form_email"]/div[1]/div[1]/label[2]/div/ins').click()

        # Andamentos no corpo do e-mail e em anexo
        driver.find_element(By.XPATH, '//*[@id="form_email"]/div[10]/div/label[2]/div/ins').click()

        # Excel - Sintético
        driver.find_element(By.XPATH, '//*[@id="form_email"]/div[11]/div/label[3]/div/ins').click()

        # Preenche campo "Mensagem extra"
        driver.find_element(By.XPATH, '//*[@id="form_email"]/div[12]/div/div[2]/div[6]/p').send_keys(self.default_text)

    @staticmethod
    def confirm_send(driver):
        """Pressiona o botão send

        Parameters:
            driver(object): Selenium Webdriver
            """
        driver.find_element(By.XPATH, '//*[@id="form_email"]/div[13]/div/div/button').click()

    @staticmethod
    def check_nothing_new(driver):
        """Verifica se a mensagem de nenhuma movimentação aparece.

        Parameters:
            driver(object): Selenium Webdriver
            """
        try:
            driver.find_element(By.XPATH, '//*[@id="gritter-item-1"]/div[2]/div[1]/p')
        except NoSuchElementException:
            return False
        return True

    @staticmethod
    def check_loading(driver):
        """Verifica se está carregando o envio

                Parameters:
            driver(object): Selenium Webdriver
            """
        try:
            if driver.find_element(By.XPATH, '//*[@class="modal fade in"]').is_displayed():
                return True
            else:
                return False
        except NoSuchElementException:
            return False

    @staticmethod
    def check_report(driver):
        """Verifica se janela com relatório está disponível

        Parameters:
            driver(object): Selenium Webdriver
            """
        try:
            driver.find_element(By.XPATH, '//*[@id="loader"]/div/div').click()
        except ElementNotInteractableException:
            return False
        return True

    @staticmethod
    def check_report_available(driver):
        """Verifica se a caixa com relatório está disponível.

        Parameters:
            driver(object): Selenium Webdriver
            """
        try:
            driver.find_element(By.XPATH, '//*[@id="modalRelatorio"]/div/div[2]')
        except NoSuchElementException:
            return True
        return False

    def get_text(self, driver):
        """Puxa as informações do relatório.

        Parameters:
            driver(object): Selenium Webdriver
            """
        driver.find_element(By.XPATH, '//*[@id="modalRelatorio"]/div/div[2]').click()
        info_panel = driver.find_element(By.XPATH, '//*[@id="filtroRelatorio"]').text
        report = driver.find_element(By.XPATH, '//*[@id="modalRelatorio"]/div/div[2]/div[2]').text
        return info_panel, report

    def run_processes(self, company, user, password, codes):
        """Roda o processo.

        Parameters:
            company (str):  Empresa - Parte 1 do Login
            user (str):     Usuário - Parte 2 do login
            password (str): Senha - Parte 3 do login
            codes (list):  Códigos aos quais serão enviados os andamentos
        """
        try:  # Tenta fazer a parte importante. Se der erro, para o programa e não envia nada, se não der erro, ok.
            nothing_new = False
            driver = self.init_driver()
            self.login_solucionare(driver, self.url, company, user, password)
            self.email_send(driver, codes)
        except Exception as e:
            print(e)
            return

        self.confirm_send(driver)

        # Aguarda o fim do carregamento
        while self.check_loading(driver):
            print(f"Carregando envios - {datetime.now().strftime('%H:%M:%S')}")
            if self.check_nothing_new(driver):
                nothing_new = driver.find_element(By.XPATH, '//*[@id="gritter-item-1"]/div[2]/div[1]/p').text

        # Verifica se a janelinha informando que não houveram novos andamentos apareceu
        if nothing_new == "Nenhum email enviado":
            data = f"Não há movimentações para os códigos {codes}"
            print(data)

        else:
            # Aguarda a janela do relatório ficar disponível
            while self.check_report(driver):
                print(f"Carregando relatório - {datetime.now().strftime('%H:%M:%S')}")

            header, body = self.get_text(driver)
            data = f"{header}\n\n{body}"
            print(data)

        with open(f"{codes[0]}-{datetime.today().strftime('%d.%m.%Y')}.txt", "w", encoding='UTF8', newline='') as f:
            f.write(data)
        f.close()

        input("Fechar janela: ")