from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from config import INVOICE_FOLDER
from config import CHROMEDRIVER_PATH
import loginulator_config as config
import time

class Loginulator:

    def __init__(self, email):
        self.driver = None # Initialize the driver to None

        # Dictionaey where keys are vendor emails and values are functions for specific vendor
        potential_logins = {config.ADP: self.ADP, config.AMEX: self.AMEX, config.DELTA: self.DELTA} 

        # Check if email is in dictionary and run the corresponding function
        for email_address in potential_logins:
            if email_address in email:
                self.setup_driver(potential_logins[email_address])

    def setup_driver(self, email_function):
        # Configure Chrome options
        chrome_options = Options()

        # Configure Chrome options for AMEX
        if email_function == self.AMEX:
            #chrome_options.add_argument("--headless")  # Disable GUI
            #chrome_options.add_argument("--disable-dev-shm-usage")  # Disable shared memory usage
            chrome_options.add_experimental_option( # Set download directory
                "prefs", {"download.default_directory": INVOICE_FOLDER}
            )

        # Initialize the Chrome driver
        self.driver = webdriver.Chrome(service=Service(CHROMEDRIVER_PATH), options=chrome_options)
        
        # Run script for specific vendor
        email_function() 

    def login(self, username, password, login_url, username_field_id, password_field_id):
        # Go to the login page.
        self.driver.get(login_url)

        # Wait for the username field to be present.
        username_field = WebDriverWait(self.driver, 20).until(EC.presence_of_element_located((By.ID, username_field_id)))
        
        # Enter the username.
        username_field.send_keys(username)

        # Simulate pressing Enter.
        username_field.send_keys(Keys.RETURN)

        # Wait for the password field to be present.
        password_field = WebDriverWait(self.driver, 20).until(EC.presence_of_element_located((By.ID, password_field_id)))
        
        # Enter the password.
        password_field.send_keys(password)

        # Submit the login form.
        password_field.submit()

    def ADP(self):
        try:
            self.login(self.driver, config.ADP_USER, config.ADP_PASS, config.ADP_LINK, "login-form_username", "login-form_password")

        except Exception as e:
            print(f"{str(e)}")
        
        finally:
            self.driver.quit()

    def AMEX(self):
        try:
            def download_statements():
                # Wait for the page to load
                WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, "//h1[contains(text(), 'Download PDF Statements')]")))
                
                # Simulate clicking on most recent statement to download
                statement_button = self.driver.find_element(By.XPATH, "//a[contains(@title, 'Download PDF Statements')]")
                statement_url = statement_button.get_attribute("href")
                self.driver.get(statement_url)
                time.sleep(1) 

            def switch_account(account_number):
                # Get the account switcher button
                account_switcher_button = self.driver.find_element(By.XPATH, "//button[contains(@class, 'axp-account-switcher__accountSwitcher__togglerButton')]")
                account_switcher_button.click()

                time.sleep(1)

                # Wait for the account switcher menu to load
                WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, "//div[@class='css-ffrkq2']")))

                # Locate and click on the specific credit card 
                credit_card_link = self.driver.find_element(By.XPATH, f"//button[contains(text(), '{account_number}')]")
                credit_card_link.click()
                
                # Simulate clicking on the statements page
                statements_link = self.driver.find_element(By.XPATH, "//a[@href='/activity?inav=myca_statements']")
                statements_link.click()

                # Wait for the statements page to load
                WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, "//a[@href='/activity/statements']")))

                # Wait for the statements page to load
                WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, "//a[@href='/activity/statements']")))

                # Click on the "All PDF Billing Statements" button
                all_statements_button = self.driver.find_element(By.XPATH, "//a[@href='/activity/statements']")
                all_statements_button.click()
                
            # Login to the default account
            self.login(config.AMEX_USER, config.AMEX_PASS, config.AMEX_LINK, "eliloUserID", "eliloPassword")

            # Wait for the page to load
            WebDriverWait(self.driver, 20).until(EC.url_changes(config.AMEX_LINK))

            # Go to the statements page
            self.driver.get(config.AMEX_STMT_LINK)

            # Download the statements from the default account
            download_statements()

            # Switch to other account and download the statements
            switch_account(config.AMEX_CC1)
            download_statements()

            # Switch to other account and download the statements
            switch_account(config.AMEX_CC2)
            download_statements()

            # Wait for the download to complete
            time.sleep(5)

        except Exception as e:
            print(f"{str(e)}")

        finally:
            self.driver.quit()

    def DELTA(self):
        try:
            # Login to the default account
            self.login(config.DELTA_USER, config.DELTA_PASS, config.DELTA_LINK, "username", "password")

            # Wait for the page to load
            WebDriverWait(self.driver, 20).until(EC.url_changes(config.DELTA_LINK))

            # Go to the activites page
            self.driver.get(config.DELTA_ACT_LINK)

            # Simulate clicking on most recent invoice to download
            statement_button = WebDriverWait(self.driver, 20).until(EC.presence_of_element_located((By.XPATH, "//a[contains(@href, '/ebillpay/my-invoices/')]")))
            statement_button.click()

            # Wait for page to load and press drop down button
            drop_down = WebDriverWait(self.driver, 20).until(EC.presence_of_element_located((By.XPATH, "//button[contains(@class, 'MuiButtonBase-root MuiButton-root')]")))
            drop_down.click()

            # Press the download button
            download_button = self.driver.find_element(By.XPATH, "//button[contains(text(), 'Download invoice PDF')]")
            download_button.click()
            
            # Wait for the download to complete
            time.sleep(15)

        except Exception as e:
            print(f"{str(e)}")
        
        finally:
            self.driver.quit()

if __name__ == "__main__":
    loginulator = Loginulator("delta")