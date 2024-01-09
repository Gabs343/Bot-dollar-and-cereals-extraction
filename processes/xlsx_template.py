import os
import pandas as pd
from datetime import datetime
from exceptions import *

class XLSXTemplateProcess:
    def __init__(self) -> None:
        self.__output_folder_path: str = self.__get_folder_path(folder='output')
        self.__template_folder_path: str = self.__get_folder_path(folder='template')
        self.__create_folders(folders=[self.__output_folder_path,
                                       self.__template_folder_path])
        
        self.__template_file_name: str = 'template-data.xlsx'
        self.__output_file_name: str = f'{datetime.now().strftime("%d.%m.%Y_%H%M%S")}.xlsx'
        
        self.__template_file_path: str = f'{self.__template_folder_path}\\{self.__template_file_name}'
        
    @property
    def file_path(self) -> str:
        return f'{self.__output_folder_path}\\{self.__output_file_name}'
    
    def __get_folder_path(self, folder: str) -> str:
        path: str = os.path.abspath(__file__).split("\\")[0:-2]
        path = '\\'.join(path)
        return f'{path}\\{folder}'
    
    def __create_folders(self, folders: list[str]) -> None:
        for folder in folders:
            if(not os.path.exists(folder)):
                os.makedirs(folder)
    
    def __columns_exists(self, columns: list[str]) -> bool:
        columns_needed = ('Name', 'Web site', 'Data')
        for column in columns:
            if(not column in columns_needed):
                return False
        return True
        
    def create_output_with_data(self, data: dict) -> None:
        if(os.path.exists(self.__template_file_path)):
            xlsx_data: dict = pd.read_excel(self.__template_file_path).to_dict()
            
            if(self.__columns_exists(columns=xlsx_data.keys())):
                row: int = 0
                for sub_dict in data.values():
                   for key, value in sub_dict.items():
                        xlsx_data['Data'][row] = f'Price {key}: ARS $ {value}'
                        row +=1

                if(os.path.exists(self.__output_folder_path)):
                    dataframe = pd.DataFrame.from_dict(xlsx_data, orient="index").T
                    dataframe.to_excel(self.file_path)
                    
                else: raise PathNotExistException('the output folder path not exist')
            else: raise MissingColumnsException('the template is missing some columns')      
        else: raise PathNotExistException('the template file path not exist')