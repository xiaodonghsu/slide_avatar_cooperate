import win32com.client
import asyncio
import json
import os
import websockets
import time

class Config():
    def __init__(self, work_mode="manual", server_host="localhost", websocket_port=8765):
        '''
        work_mode: 工作模式, manual 手动, collaboration 协同, auto 自动
        server_host: WebSocket 服务器监听地址
        websocket_port: WebSocket 服务器监听端口
        '''
        self.__CONFIG_FILE = "config.json"
        if not os.path.exists(self.__CONFIG_FILE):
            raise FileNotFoundError(f"Configuration file '{self.__CONFIG_FILE}' not found.")
        self.config = {}
        self.__last_update_time = None
        self.update()
        self.isUpdated = False

    def update(self):
        self.isUpdated = False
        if not self.__last_update_time == os.path.getmtime(self.__CONFIG_FILE):
            print("Configuration file updated, reloading...")
            with open(self.__CONFIG_FILE, "r") as f:
                self.config = json.load(f)
            self.__last_update_time = os.path.getmtime(self.__CONFIG_FILE)
            self.isUpdated = True
        if self.isUpdated:
            print("Current Configuration:", self.config)
            if self.config["work_mode"] == "manual":
                print('讲解员模式, 检测ppt状态，不自动播放')
            elif self.config["work_mode"] == "collaboration":
                print("协作模式, 讲解员决定内容播放")
            elif self.config["work_mode"] == "auto":
                print("自动模式, ppt 启动后自动播放")
            else:
                print("未知模式:", self.config["work_mode"], "按讲解员模式处理")
                self.config["work_mode"] = "manual"

class PowerPointMonitor():
    def __init__(self):
        self._ppt_app_list = ["PowerPoint.Application", "Kwpp.Application"]
        self.ppt_app_name = None
        self.ppt_app = None
        self.connect_powerpoint()
        self.slide_show_active = False
        self.presentation_name = None

    def connect_powerpoint(self):
        for app_name in self._ppt_app_list:
            try:
                self.ppt_app = win32com.client.GetActiveObject(app_name)
                if self.ppt_app:
                    self.ppt_app_name = app_name
                    print(f"连接 ppt app {app_name} 成功!")
                    break
            except Exception as e:
                # print(e)
                self.ppt_app = None

    def get_presentation_name(self):
        '''
        获取当前活动的演示文稿名称
        '''
        if self.get_presentations_count()>0:
            try:
                presentation = self.ppt_app.ActivePresentation
                return presentation.Name
            except:
                return None

    def get_presentations_count(self):
        '''
        根据 PowerPoint 应用程序对象判断是否连接
        -1: 未连接
        0: 已连接但无打开的演示文稿
        n: 已连接且打开了 n 个演示文稿
        '''
        present_count = -1
        if self.ppt_app:
            try:
                present_count = self.ppt_app.Presentations.Count
            except:
                present_count = -1
        return present_count

    def get_edit_slide_index(self):
        '''
        获取编辑的幻灯片编号
        '''
        if self.get_presentations_count()>0:
            try:
                presentation = self.ppt_app.ActivePresentation
                slide = presentation.Windows(1).View.Slide
                return slide.SlideIndex
            except:
                return -1

    def get_present_slide_index(self):
        '''
        获取播放的幻灯片编号
        '''
        if self.get_presentations_count()>0:
            try:
                presentation = self.ppt_app.ActivePresentation
                slide = presentation.SlideShowWindow.View.Slide
                return slide.SlideIndex
            except:
                return -1

    def get_current_status(self):
        '''
        获取当前幻灯片的播放状态和编号, 返回 (文件数量, 编辑编号, 放编号)
        编号为 -1 无效
        '''
        # 检查当前连接状态
        present_count = self.get_presentations_count()
        # 如果没有连接 PowerPoint，则尝试连接
        if present_count == -1:
            self.connect_powerpoint()
        present_count = self.get_presentations_count()
        # 根据当前连接状态返回结果
        if present_count == -1:
            return present_count, -1, -1
        if present_count == 0:
            return present_count, -1, -1
        if present_count > 0:
            # 获取当前编辑状态的幻灯片编号
            edit_index = -1
            try:
                presentation = self.ppt_app.ActivePresentation
                slide = presentation.Windows(1).View.Slide
                edit_index = slide.SlideIndex
            except:
                pass

            # 则获取当前幻灯片编号
            present_index = -1
            try:
                presentation = self.ppt_app.ActivePresentation
                slide = presentation.SlideShowWindow.View.Slide
                present_index = slide.SlideIndex
            except:
                pass
            return present_count, edit_index, present_index

async def broadcast_slide_change():

    # 首次启动使用当前的配置作为基础配置
    print("读取配置文件...")
    cfg = Config()
    previous_config = cfg.config.copy()
    previous_edit_slide_index = -1
    previous_present_slide_index = -1

    # 初始化 PowerPoint 监控器
    ppt_monitor = PowerPointMonitor()
    ppt_name = ppt_monitor.get_presentation_name()
    present_count, previous_edit_slide_index, previous_present_slide_index = ppt_monitor.get_current_status()
    if ppt_name is None:
        print("No PPT opened")
    else:
        print("PPT应用名称:", ppt_monitor.ppt_app_name, 
              "PPT名称:", ppt_name) 
        print("当前编辑页面:", previous_edit_slide_index,
              "当前播放页面:", previous_present_slide_index,
              "总页面数:", present_count)

    # 启动 WebSocket 服务器
    async with websockets.serve(handler.handler, cfg.config["server_host"], cfg.config["websocket_port"]):
        print(f'WebSocket server started at ws://{cfg.config["server_host"]}:{cfg.config["websocket_port"]}')
        while True:
            # 监测配置文件更新
            # 如果文件更新, 需要分析具体的变化:
            # 从 auto 或 collaboration 切换到 manual, 需要发送停止播放的消息
            cfg.update()
            if cfg.isUpdated:
                if cfg.config["work_mode"] == "manual":
                    print("Switched to manual mode.") 
                    if previous_config["work_mode"] != cfg.config["work_mode"]:
                        print(time.strftime("%H-%M-%S"), f'Switched to {cfg.config["work_mode"]} mode.')
                        # 发送空的播放列表以停止播放
                        message = json.dumps({
                            "tasks": "playlist",
                            "playlist": []
                        })
                        await handler.send_to_clients(message)
                        pass

                # 更新后的配置处理检查avatar_status变化
                if previous_config["avatar_status"] != cfg.config["avatar_status"]:
                    print("发送任务")
                    message = json.dumps({
                            "tasks": cfg.config["avatar_status"]
                        })
                    await handler.send_to_clients(message)
            # 更新 previous_config
            previous_config = cfg.config.copy()
 
            # 监测 PPT 状态变化
            present_count, edit_index, present_index = ppt_monitor.get_current_status()
            if previous_present_slide_index != present_index or previous_edit_slide_index != edit_index:
                print(f"页面总数: {present_count}, 编辑页面: {edit_index}, 播放页面: {present_index}")
            
            if present_count > 0:
                new_slide_index = -1
                if (present_index != previous_present_slide_index):
                    new_slide_index = present_index
                else:
                    if (edit_index != previous_edit_slide_index):
                        new_slide_index = edit_index
                # print(new_slide_index)
                if new_slide_index != -1:
                    print(time.strftime("%H-%M-%S"), f"Status: {present_count}, Current edit slide: {edit_index}, Current Present slide: {present_index}")
                    message = json.dumps({
                        "tasks": "playlist",
                        "playlist": [
                            {"video": f"../assets/videos/video{new_slide_index}.webm", "loop": 1},
                            {"video": "../assets/videos/idle.webm", "loop": -1}
                        ]
                    })
                    await handler.send_to_clients(message)
            await asyncio.sleep(0.5)

            previous_present_slide_index = present_index
            previous_edit_slide_index = edit_index


            # 手动模式下不监测 PPT 状态
            # if cfg.config["work_mode"] == "manual":
            #     # 如果之前不是 manual 模式，则发送停止播放的消息
            #     previous_config = cfg.config.copy()
            #     await asyncio.sleep(1)
            #     continue

            # # 协作模式
            # if cfg.config["work_mode"] == "collaboration":
            #     if previous_config["work_mode"] != cfg.config["work_mode"]:
            #         print(time.strftime("%H-%M-%S"), "Switched to collaboration work_mode.")
            #         # 发送空的播放列表以停止播放
            #         message = json.dumps({
            #             "tasks": "playlist",
            #             "playlist": []
            #         })
            #         await handler.send_to_clients(message)

            
            # if cfg.config["work_mode"] == "auto":
            #     # present_count, edit_index, present_index = ppt_monitor.get_current_status()
            #     # print(f"{present_count}, {edit_index}, {present_index}")
            #     pass


# 管理所有连接的客户端
class handler:
    clients = set()
    
    @classmethod
    async def handler(cls, websocket, path=None):
        cls.clients.add(websocket)
        client_addr = None
        try:
            # 尝试获取客户端地址以便打印日志
            try:
                client_addr = websocket.remote_address
            except Exception:
                client_addr = None
            print(f"Client connected: {client_addr}")
            try:
                async for message in websocket:
                    try:
                        print(f"Received from {client_addr}: {message}")
                    except Exception as e:
                        print("Error printing received message:", e)
                    # 尝试解析为 JSON 并打印解析后的内容
                    try:
                        parsed = json.loads(message)
                        print("Parsed message:", parsed)
                    except Exception:
                        # 非 JSON 消息则忽略解析错误
                        pass
            except websockets.exceptions.ConnectionClosed:
                # 连接被客户端正常或异常关闭
                pass
        finally:
            cls.clients.remove(websocket)
            print(f"Client disconnected: {client_addr}")
    
    @classmethod
    async def send_to_clients(cls, message):
        print(f"Broadcasting message to {len(cls.clients)} clients: {message}")
        if cls.clients:
            # 原版
            # await asyncio.wait([client.send(message) for client in cls.clients])
            # 某些版本下会出现
            # Passing coroutines coroutines is forbidden use tasks explicitly
            # 的提示, 需要先转换为任务
            tasks = [asyncio.create_task(client.send(message)) for client in cls.clients]
            await asyncio.wait(tasks)


if __name__ == "__main__":
    asyncio.run(broadcast_slide_change())
