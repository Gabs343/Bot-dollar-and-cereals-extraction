from datetime import datetime, timedelta

class SettingService: 
    def __init__(self):
        self.__settings: dict = {}

    @property
    def settings(self) -> dict:
        return self.__settings
        
    @settings.setter
    def settings(self, setting) -> None:
        self.__settings = setting
    
class BcrSetting(SettingService):
    
    def __init__(self, bot_name: str) -> None:
        super().__init__()
        self.settings = self.get_new_settings()
        
    def get_new_settings(self) -> dict:
        setting: dict = {}
        end_date: datetime = datetime.now()
        start_date: datetime = end_date - timedelta(days=7)
        setting['data_1'] = {'product': 'Maíz', 
                            'price': 'Precios Pizarra', 
                            'startDate': start_date.strftime('%m%d%Y'), 
                            'endDate': end_date.strftime('%m%d%Y')}
        
        setting['data_2'] = {'product': 'Soja', 
                            'price': 'Precios Pizarra', 
                            'startDate': start_date.strftime('%m%d%Y'), 
                            'endDate': end_date.strftime('%m%d%Y')}

        return setting
        
    def __str__(self) -> str:
        return type(self).__name__