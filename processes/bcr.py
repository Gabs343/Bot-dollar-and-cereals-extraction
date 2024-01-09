from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement

class BCRProcess:
    def __init__(self) -> None:
        self.__url: str = 'https://www.cac.bcr.com.ar/es/precios-de-pizarra/consultas'
        
    def open(self) -> None:
        self.driver = webdriver.Edge()
        self.driver.maximize_window()
        self.driver.get(self.__url)
        self.driver.implicitly_wait(0.5)
        
    def close(self) -> None:
        self.driver.quit()
        
    def set_product(self, product: str) -> None:
        self.driver.find_element(by=By.CSS_SELECTOR, value='[data-id="edit-product"]').click()
        elements: WebElement = self.driver.find_elements(by=By.CLASS_NAME, value='text')
        for element in elements:
            if(element.get_attribute('innerHTML') == product):
                element.click()
                break
    
    def set_price(self, price: str) -> None:
        self.driver.find_element(by=By.CSS_SELECTOR, value='[data-id="edit-type"]').click()
        elements: WebElement = self.driver.find_elements(by=By.CLASS_NAME, value='text')
        for element in elements:
            if(element.get_attribute('innerHTML') == price):
                element.click()
                break
            
    def set_date(self, start: str, end: str) -> None:
        start_date_input: WebElement = self.driver.find_element(by=By.ID, value='edit-date-start')
        end_date_input: WebElement = self.driver.find_element(by=By.ID, value='edit-date-end')
        start_date_input.send_keys(start)
        end_date_input.send_keys(end)
        
    def click_filter(self) -> None:
        self.driver.find_element(by=By.ID, value='edit-submit').click()
        
    def clean_filter(self) -> None:
        self.driver.find_element(by=By.CLASS_NAME, value='btn-reset').click()
        
    def get_last_price_for_product(self) -> str:
        last_td_element: WebElement = self.driver.find_elements(by=By.TAG_NAME, value='td')[-1]
        return last_td_element.get_attribute('innerHTML').strip()[1:]