import os
import json

class Config():
    def __init__(self):
        '''
        work_mode: 工作模式, manual 手动, collaboration 协同, auto 自动
        server_host: WebSocket 服务器监听地址
        websocket_port: WebSocket 服务器监听端口
        '''
        self.__CONFIG_FILE = "config.json"
        self.__DEFAULT_CONFIG_FILE = "config.json.default"
        if not os.path.exists(self.__CONFIG_FILE):
            # 默认配置文件复制为config.json
            with open(self.__DEFAULT_CONFIG_FILE, 'rb') as f:
                data = f.read()
            with open(self.__CONFIG_FILE, 'wb') as f:
                f.write(data)
        self.config = {}
        self.__last_load_time = None
        self.isFresh = False
        self.__config_command_response_name__ = "avatar_command_response"
        self.__config_work_mode_response_name__ = "work_mode_response"

    def load_config(self):
        with open(self.__CONFIG_FILE, "r", encoding='utf-8') as f:
            self.config = json.load(f)
        self.__last_load_time = os.path.getmtime(self.__CONFIG_FILE)

    def dump_config(self):
        with open(self.__CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(self.config, f, ensure_ascii=False, indent=4)

    def fresh(self):
        '''
        检查配置文件是否有更新, 如果有更新, 重新加载配置文件
        '''
        self.isFresh = False
        if not self.__last_load_time == os.path.getmtime(self.__CONFIG_FILE):
            self.load_config()
            self.isFresh = True
        else:
            self.isFresh = False

    def update_work_mode_response(self, value=None):
        self.load_config()
        if value is None:
            value = {"result": "success", "reason": ""}
        self.config[self.__config_work_mode_response_name__] = value
        self.dump_config()

    def update_avatar_command_response(self, value=None):
        self.load_config()
        if value is None:
            value = {"result": "success", "reason": ""}
        self.config[self.__config_command_response_name__] = value
        self.dump_config()
