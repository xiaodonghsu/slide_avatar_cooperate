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
        self.isUpdated = False

    def load_config(self):
        with open(self.__CONFIG_FILE, "r", encoding='utf-8') as f:
            self.config = json.load(f)
        self.__last_load_time = os.path.getmtime(self.__CONFIG_FILE)

    def dump_config(self):
        with open(self.__CONFIG_FILE, "w") as f:
            json.dump(self.config, f, ensure_ascii=False, indent=4)

    def update(self):
        self.isUpdated = False
        if not self.__last_load_time == os.path.getmtime(self.__CONFIG_FILE):
            self.load_config()
            self.isUpdated = True
