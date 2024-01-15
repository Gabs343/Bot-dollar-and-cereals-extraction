import time
import sys
import win32

from logs import *
from settings import *
from exceptions import *
from processes.ambito import AmbitoProcess
from processes.bcr import BCRProcess
from processes.xlsx_template import XLSXTemplateProcess
  
class Main:
    __settings_services_classes: tuple = (BcrSetting, )
    __logs_services_classes: tuple = (LogTxt, LogXlsx, LogVideo)
    __settings_services: list[SettingService] = []
    __logs_services: list[LogService] = []
    __bot_name: str = "Web-Extraction"
    __status: str = "READY"
    __status_callback = None
    __had_error: bool = False
    
    def __init__(self) -> None:
        self.__settings_services = self.__get_settings_services()
        
    @property   
    def settings_services(self) -> list[SettingService]:
        return self.__settings_services
    
    @property   
    def logs_services(self) -> list[LogService]:
        return self.__logs_services
    
    @property   
    def bot_name(self) -> str:
        return self.__bot_name
    
    @property   
    def status(self) -> str:
        return self.__status
    
    @property   
    def status_callback(self) -> str:
        return self.__status_callback
    
    @status_callback.setter
    def status_callback(self, callback) -> None:
        self.__status_callback = callback
        
    def start(self, *args) -> None:
        try:
            self.__execution_begun()

            #Your code goes here
            data: dict = {}
            data = {**data, **self.do_ambito_proccess()}
            data = {**data, **self.do_bcr_proccess()}
            output_path: str = self.do_xlsx_proccess(data=data)
            print(data)
            ####################
            
            self.__execution_completed()
            self.do_sent_mail_proccess(attachments=[output_path])
        except Exception as e: 
            raise Exception(e)
        
    def do_ambito_proccess(self) -> dict:
        try:
            logXlsx: LogXlsx = self.__execute_action(function=self.__get_log_service, log_type=LogXlsx)
            logTxt: LogTxt = self.__execute_action(function=self.__get_log_service, log_type=LogTxt)
            
            self.__execute_action(function=logTxt.write_info, message='Ambito process has begun')
            
            data: dict = {}
            data['ambito'] = dict.fromkeys(['dollar'], 'Error')
            
            page: AmbitoProcess = AmbitoProcess()
            self.__execute_action(function=page.open)
            data['ambito']['dollar'] = self.__execute_action(function=page.extract_dollar)
            self.__execute_action(function=page.close)
            
            self.__execute_action(function=logXlsx.write_info, message='Ambito Process')
            self.__execute_action(function=logTxt.write_info, message='Ambito process completed')
        
        except Exception as e:
            self.__execute_action(function=logXlsx.write_error, message='Ambito Process', detail='There was an error in the ambito page')
            self.__execute_action(function=logTxt.write_error, message='Ambito process not completed', detail=e)
            self.__had_error = True
            
        finally: return data
       
    def do_bcr_proccess(self) -> dict:
        try:
            logXlsx: LogXlsx = self.__execute_action(function=self.__get_log_service, log_type=LogXlsx)
            logTxt: LogTxt = self.__execute_action(function=self.__get_log_service, log_type=LogTxt)
            
            self.__execute_action(function=logTxt.write_info, message='Bcr process has begun')
            
            data: dict = {}
            
            bcr_setting_service: BcrSetting =  self.__execute_action(function=self.__get_setting_service, setting_type=BcrSetting)
            products: list[str] = [sub_dict['product'] for sub_dict in bcr_setting_service.settings.values()]
            data['bcr'] = dict.fromkeys(products, 'Error')
  
            page: BCRProcess = BCRProcess()
            self.__execute_action(function=page.open)
            
            for request in bcr_setting_service.settings.values():
                try:
                    self.__execute_action(function=page.set_product, product=request['product'])
                    self.__execute_action(function=page.set_price, price=request['price'])
                    self.__execute_action(function=page.set_date, start=request['startDate'], end=request['endDate'])
                    self.__execute_action(function=page.click_filter)
                    data['bcr'][request['product']] = self.__execute_action(function=page.get_last_price_for_product)
                    self.__execute_action(function=page.clean_filter)
                except Exception as e:
                    self.__execute_action(function=logXlsx.write_error, message='Bcr Process', detail=f'Error while getting the price for {request["product"]}')
                    self.__execute_action(function=logTxt.write_error, message=f'Error while getting the price for {request["product"]}\n', detail=e)
                    self.__had_error = True
            
            self.__execute_action(function=page.close)
            
            self.__execute_action(function=logXlsx.write_info, message='Bcr Process')
            self.__execute_action(function=logTxt.write_info, message='Bcr process completed')
            
        except Exception as e:
            self.__execute_action(function=logXlsx.write_error, message='Bcr Process', detail='There was an error in the bcr page')
            self.__execute_action(function=logTxt.write_error, message='Bcr process not completed', detail=e)
            self.__had_error = True
            
        finally: return data
    
    def do_xlsx_proccess(self, data: dict) -> str:
        try:
            logXlsx: LogXlsx = self.__execute_action(function=self.__get_log_service, log_type=LogXlsx)
            logTxt: LogTxt = self.__execute_action(function=self.__get_log_service, log_type=LogTxt)
            
            template: XLSXTemplateProcess = XLSXTemplateProcess()
            self.__execute_action(function=template.create_output_with_data, data=data)
            
            self.__execute_action(function=logXlsx.write_info, message='Xlsx template Process')
            self.__execute_action(function=logTxt.write_info, message='Xlsx template completed')
            
            return template.file_path
            
        except PathNotExistException as e:
            self.__execute_action(function=logXlsx.write_error, message='Xlsx template Process', detail=e)
            self.__execute_action(function=logTxt.write_error, message='Xlsx template process not completed', detail=e)
            self.__had_error = True
            
        except MissingColumnsException as e:
            self.__execute_action(function=logXlsx.write_error, message='Xlsx template Process', detail=e)
            self.__execute_action(function=logTxt.write_error, message='Xlsx template process not completed', detail=e)
            self.__had_error = True
            
        except Exception as e:
            self.__execute_action(function=logXlsx.write_error, message='Xlsx template Process', detail='There was a general error while creating the output data')
            self.__execute_action(function=logTxt.write_error, message='Xlsx template process not completed', detail=e)
            self.__had_error = True
    
    def do_sent_mail_proccess(self, attachments: list[str]=None) -> None:
        logs_paths = [log.file_path for log in self.__logs_services]
        attachments.extend(logs_paths)
        if(self.__had_error):
            self.send_outlook(to_mail='', 
                              subject='Bot execution completed with errors',
                              body='', attachments=attachments)
        else:
            self.send_outlook(to_mail='', 
                              subject='Bot execution completed without errors', 
                              body='', 
                              attachments=attachments)
              
    def send_outlook(to_mail: str, subject: str, body: str, attachments: list[str]=None) -> None:
        outlook = win32.Dispatch("outlook.application")
        mail = outlook.CreateItem(0)
        mail.To = to_mail
        mail.Subject = subject
        mail.Body = body

        if(attachments != None):
            for file in attachments:
                if(os.path.exists(file)):
                    mail.Attachments.Add(file)

        mail.Send()
            
    def pause(self) -> None:
        self.__notify_status(new_status='PAUSED')
            
    def unpause(self) -> None:
        self.__notify_status(new_status='RUNNING')
        
    def stop(self) -> None:
        self.__notify_status(new_status='CLOSING BOT')
        
    def __execution_begun(self) -> None:
        log_name: str = datetime.now().strftime("%d.%m.%Y_%H%M%S")
        self.__logs_services = [log(name=log_name) for log in self.__logs_services_classes]
        logXlsx: LogXlsx = self.__get_log_service(log_type=LogXlsx)
        logXlsx.write_info(message=f'The Bot has begun')
        self.__notify_status(new_status="RUNNING")
             
    def __execution_completed(self) -> None:
        self.__notify_status(new_status="READY")
        logXlsx: LogXlsx = self.__get_log_service(log_type=LogXlsx)
        if(self.__had_error):
            logXlsx.write_error(message=f'The Bot has ended')
        else:
            logXlsx.write_info(message=f'The Bot has ended')
        self.__close_logs()
           
    def __notify_status(self, new_status: str) -> None:
        self.__status = new_status
        logTxt: LogTxt = self.__get_log_service(log_type=LogTxt)
        logTxt.write_info(message=f'Bot {new_status}')
        if self.__status_callback:
            self.__status_callback(new_status)
        
    def __get_log_service(self, log_type: LogService) -> LogService:
        try: return next(log for log in self.__logs_services if isinstance(log, log_type))
        except StopIteration:
            raise ServiceNotFound(f'The log service of type {log_type}, cannot be found')
    
    def __get_setting_service(self, setting_type: SettingService) -> SettingService:
        try: return next(service for service in self.__settings_services if isinstance(service, setting_type))
        except StopIteration:
            raise ServiceNotFound(f'The setting service of type {setting_type}, cannot be found')
        
    def __execute_action(self, function, **kwargs):
        logTxt: LogTxt = self.__get_log_service(log_type=LogTxt)
        while self.__status == 'PAUSED':
            if(self.__status=='RUNNING'):
                break
        return logTxt.write_and_execute(function, **kwargs)
    
    def __close_logs(self) -> None:
        for log in self.__logs_services:
            log.close()
    
    def __get_settings_services(self) -> list[SettingService]:
        return [service(bot_name=self.__bot_name) for service in self.__settings_services_classes]
                                
if __name__ == "__main__":
    st = time.time()
    main = Main()
    main.start(sys.argv[1:])
    et = time.time()
    elapsed_time = et - st
    print('Execution time:', elapsed_time, 'seconds')
        
    
