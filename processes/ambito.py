from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement

class AmbitoProcess:
    def __init__(self) -> None:
        self.__url: str = 'https://www.ambito.com/contenidos/dolar-cl.html'
        
    def open(self) -> None:
        self.driver = webdriver.Edge()
        self.driver.maximize_window()
        self.driver.get(self.__url)
        self.driver.implicitly_wait(0.5)
        
    def close(self) -> None:
        self.driver.quit()
        
    def extract_dollar(self) -> str:
        span_element: WebElement = self.driver.find_element(by=By.CLASS_NAME, value='data-valor')
        return span_element.get_attribute('innerHTML').strip()