import win32com.client
import asyncio
import json
import os
import websockets
import time
import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path

# Logger setup: write logs to project root `monitor.log`
if not os.path.exists("../log"):
    os.mkdir("log")
LOG_FILE = os.path.join("../log", "monitor.log")
logger = logging.getLogger("monitor_service")
logger.setLevel(logging.DEBUG)
if not logger.handlers:
    fh = RotatingFileHandler(LOG_FILE, maxBytes=5 * 1024 * 1024, backupCount=3, encoding='utf-8')
    formatter = logging.Formatter('%(asctime)s %(levelname)s [%(name)s] %(message)s')
    fh.setFormatter(formatter)
    logger.addHandler(fh)
    ch = logging.StreamHandler()
    ch.setFormatter(formatter)
    logger.addHandler(ch)

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
            logger.info("Configuration file updated, reloading...")
            with open(self.__CONFIG_FILE, "r", encoding='utf-8') as f:
                self.config = json.load(f)
            self.__last_update_time = os.path.getmtime(self.__CONFIG_FILE)
            self.isUpdated = True
        if self.isUpdated:
            logger.info("Current Configuration: %s", self.config)
            if self.config["work_mode"] == "manual":
                logger.info('讲解员模式, 检测ppt状态，不自动播放')
            elif self.config["work_mode"] == "collaboration":
                logger.info("协作模式, 讲解员决定内容播放")
            elif self.config["work_mode"] == "auto":
                logger.info("自动模式, ppt 启动后自动播放")
            else:
                logger.warning("未知模式: %s, 按讲解员模式处理", self.config["work_mode"])
                self.config["work_mode"] = "manual"

class PowerPointMonitor():
    def __init__(self):
        self._ppt_app_list = ["PowerPoint.Application", "Kwpp.Application"]
        self.ppt_app_name = None
        self.ppt_app = None
        self.slide_show_active = False
        self.presentation_name = None
        self.slide_video_list = None
        self.__assets_base_dir = "../assets"
        self.__slide_video_config = "slide_video.json"
        self.__slide_index_prefix = "slide-"
        self.__idle_video_prefix = "idle"
        # 只让 ppt_add 没有连接的信息出现一次
        self.__ppt_app_warning_flag = True

    def connect_powerpoint(self):
        for app_name in self._ppt_app_list:
            try:
                self.ppt_app = win32com.client.GetActiveObject(app_name)
                if self.ppt_app:
                    self.ppt_app_name = app_name
                    logger.info("连接PPT应用: %s 成功!", app_name)
                    break
            except Exception as e:
                self.ppt_app = None
        if self.ppt_app is None:
            if self.__ppt_app_warning_flag:
                logger.warning("没有可用的PPT应用")
                self.__ppt_app_warning_flag = False
            return
        else:
             self.__ppt_app_warning_flag = True
        
        if not self.ppt_app.Visible:
            try:
                logger.info("PPT应用不可见,设置PPT可见")
                self.ppt_app.Visible = 1
            except:
                logger.warning("设置可见性失败!")

        presentation_name = self.get_presentation_name()
        if presentation_name:
            self.update_slide_video_list(presentation_name)

    def get_presentation_name(self):
        '''
        获取当前活动的演示文稿名称
        '''
        if self.isConnected():
            try:
                presentation = self.ppt_app.ActivePresentation
                self.presentation_name = presentation.Name
                return self.presentation_name
            except:
                self.presentation_name = None
                return None
    def isConnected(self):
        return True if self.get_presentations_count() > 0 else False

    def get_presentations_count(self):
        '''
        根据 PowerPoint 应用程序对象判断是否连接
        -1: 未连接
        0: 已连接但无打开的演示文稿
        n: 已连接且打开了 n 个演示文稿
        '''
        if self.ppt_app:
            return self.ppt_app.Presentations.Count
        return -1

    def get_slides_count(self):
        if self.get_presentations_count()>0:
            try:
                presentation = self.ppt_app.ActivePresentation
                return presentation.Slides.Count
            except:
                return -1
        return -1

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

    def get_current_ppt_status(self):
        '''
        获取当前幻灯片的播放状态和编号, 返回 
        {"present_count": 打开的ppt数量,
        "present_name": ppt名字,
        "slides_count": 胶片数量,
        "edit_slide_index": 当前编辑的胶片的索引,
        "present_slide_index": 放映的胶片的索引}
        编号为 -1 无效
        '''
        # 检查当前连接状态
        current_ppt_status = {
            "present_count": -1,
            "present_name": "",
            "slides_count": -1,
            "edit_slide_index": -1,
            "present_slide_index": -1
        }
        # 如果没有连接 PowerPoint，则尝试连接
        if not self.isConnected():
            self.connect_powerpoint()
        present_count = self.get_presentations_count()
        # 根据当前连接状态返回结果
        if present_count == -1:
            return current_ppt_status
        if present_count == 0:
            current_ppt_status["present_count"] = present_count
            return current_ppt_status
        if present_count > 0:
            current_ppt_status["present_count"] = present_count
            presentation = self.ppt_app.ActivePresentation
            current_ppt_status["present_name"] = presentation.Name
            current_ppt_status["slides_count"] = presentation.Slides.Count
            current_ppt_status["edit_slide_index"] = presentation.Windows(1).View.Slide.SlideIndex
            # 则获取当前幻灯片编号
            present_slide_index = -1
            try:
                present_slide_index = presentation.SlideShowWindow.View.Slide.SlideIndex
                current_ppt_status["present_slide_index"] = present_slide_index
            except:
                pass
            return current_ppt_status

    def update_slide_video_list(self, presentation_name):
        '''
        获取当前幻灯片中的视频列表
        '''
        self.slide_video_list = None
        try:
            with open(os.path.join(self.__assets_base_dir, self.__slide_video_config), "r", encoding='utf-8') as f:
                slide_video_config = json.load(f)
        except Exception as e:
            logger.warning("处理 slide_video.json 失败: %s, 当前文件夹: %s", e, os.getcwd())
            return

        for item in slide_video_config["slide_videos"]:
            if "name" in item and "videos" in item:
                print(item["name"])
                if item["name"].endswith(presentation_name):
                    self.slide_video_list = item["videos"]
                    logger.info("当前演示文稿 %s 包含 %s 个视频", presentation_name, len(self.slide_video_list))
                    break
        if self.slide_video_list is None:
            logger.warning("当前演示文稿 %s 没有配置数字人视频,无法播放数字人!", presentation_name)
        else:
            # 确认视频是否存在,并补充完整路径
            for video_file_index in self.slide_video_list:
                video_file = os.path.join(self.__assets_base_dir, self.slide_video_list[video_file_index])
                if not os.path.exists(video_file):
                    logger.error("Video file %s does not exist", video_file)
                    self.slide_video_list[video_file_index] = None
                else:
                    self.slide_video_list[video_file_index] = video_file

    def get_slide_video_file(self, prensentation_index):
        index = self.__slide_index_prefix + str(prensentation_index)
        return self.get_video_file(index)

    def get_idle_video_file(self, prensentation_index = None):
        index = self.__idle_video_prefix
        return self.get_video_file(index)

    def get_video_file(self, prensentation_index):
        if self.slide_video_list is None:
            logger.warning("当前幻灯片没有配置数字人视频")
            return []
        if not prensentation_index in self.slide_video_list:
            logger.warning("当前幻灯片 %s 没有配置数字人视频" , prensentation_index)
            return []
        return self.slide_video_list[prensentation_index]

    def goto_page(self, dest_slide_index = 0):
        '''
        ppt 跳转到指定页面, -1, 指代上一页; 0 指代下一页; 正数指定特定页面
        如果是播放状态, 跳转播放
        如果是非播放状态, 则跳转编辑的页面
        '''
        current_ppt_status = self.get_current_ppt_status()
        logger.info(current_ppt_status)
        if current_ppt_status["slides_count"] > 0:
            if current_ppt_status["present_slide_index"] > 0:
                # 下一页
                if dest_slide_index == 0:
                    if current_ppt_status["present_slide_index"] < current_ppt_status["slides_count"]:
                        self.ppt_app.ActivePresentation.SlideShowWindow.View.Next()
                elif dest_slide_index == -1:
                    # 上一页
                    if current_ppt_status["present_slide_index"] > 1:
                        self.ppt_app.ActivePresentation.SlideShowWindow.View.Previous()
                else:
                    if dest_slide_index >= 1 and dest_slide_index <= current_ppt_status["slides_count"]:
                        self.ppt_app.ActivePresentation.SlideShowWindow.View.GotoSlide(dest_slide_index)
            elif current_ppt_status["edit_slide_index"] > 0:
                # 下一页
                if dest_slide_index == 0:
                    if current_ppt_status["edit_slide_index"] < current_ppt_status["slides_count"]:
                        self.ppt_app.ActivePresentation.Slides(current_ppt_status["edit_slide_index"] + 1).Select()
                elif dest_slide_index == -1:
                    # 上一页
                    if current_ppt_status["edit_slide_index"] > 1:
                        self.ppt_app.ActivePresentation.Slides(current_ppt_status["edit_slide_index"] - 1).Select()
                else:
                    if dest_slide_index >= 1 and current_ppt_status["edit_slide_index"] <= dest_slide_index:
                        self.ppt_app.ActivePresentation.Slides(dest_slide_index).Select()

    def goto_next_page(self):
        self.goto_page(0)

    def goto_previous_page(self):
        self.goto_page(-1)

    def start_slideshow(self):
        if self.get_presentations_count() <= 0:
            return
        try:
            self.ppt_app.ActivePresentation.SlideShowSettings.Run()
            self.slide_show_active = True
        except:
            logger.warning("开始幻灯片放映失败")


global ppt_monitor
ppt_monitor = None

def parse_event(event: dict) -> str:
    '''
    根据事件, 返回数字人的状态
    idle
    playing
    unknown
    '''
    global ppt_monitor
    if "event" in event and "type" in event and "src" in event:
        if event["event"] == "started":
            if event["type"] == "video":
                if event["src"] == ppt_monitor.get_idle_video_file():
                    return "idle"
                else:
                    return "playing"
        elif event["event"] == "finished":
            return "idle"
    return "unknown"


async def broadcast_slide_change():
    # 首次启动使用当前的配置作为基础配置
    logger.info("读取配置...")
    cfg = Config()
    previous_config = cfg.config.copy()

    logger.info("初始化 Presentation 监控...")
    global ppt_monitor
    ppt_monitor = PowerPointMonitor()
    ppt_monitor.connect_powerpoint()
    previous_ppt_status = ppt_monitor.get_current_ppt_status()
    if previous_ppt_status["present_count"] < 0:
        logger.warning("PPT应用没有打开")
    else:
        logger.info("PPT应用名称: %s", ppt_monitor.ppt_app_name)
        logger.info("初始化PPT文件名: %s", previous_ppt_status["present_name"])
        logger.info("初始化PPT状态: %s", previous_ppt_status)

    # 数字人状态: idle, playing, pause, unknown
    avatar_status = "idle"

    # 启动 WebSocket 服务器
    async with websockets.serve(handler.handler, cfg.config["server_host"], cfg.config["websocket_port"]):
        logger.info('WebSocket server started at ws://%s:%s', cfg.config["server_host"], cfg.config["websocket_port"])
        while True:
            # 监测配置文件更新
            # 检测三个参数变化 work_mode , avatar_command, avatar_status
            # avatar_command 记录需要发送到数字人的命令, 变化时, play, pause指令, 将内容变化给客户端
            # avatar_status 记录数字人播放器返回的状态事件
            # work_mode 从 auto 或 collaboration 切换到 manual, 需要发送停止播放的消息
            cfg.update()
            if cfg.isUpdated:
                # 优先处理事件,减少事件问题导致状态的变化
                if previous_config["avatar_event"] != cfg.config["avatar_event"]:
                    logger.info("数字人事件:", cfg.config["avatar_event"])
                    # 根据工作模式处理事件
                    avatar_status = parse_event(cfg.config["avatar_event"])
                    logger.info("数字人状态切换: %s", avatar_status)
                    if avatar_status == "idle":
                        if cfg.config["work_mode"] == "auto":
                            ppt_monitor.goto_next_page()
                    elif avatar_status == "playing":
                        pass
                    elif avatar_status == "unknown":
                        pass

                # 处理 work_mode 变化
                if previous_config["work_mode"] != cfg.config["work_mode"]:
                    logger.info("Switching to %s mode.", cfg.config["work_mode"])
                    if cfg.config["work_mode"] == "manual":
                        # 讲解员模式: 主要的讲解任务在讲解员。数字人不参与
                        # 发送空的播放列表以停止播放
                        message = json.dumps({
                            "tasks": "playlist",
                            "playlist": []
                        })
                        await handler.send_to_clients(message)
                    elif cfg.config["work_mode"] == "collaboration":
                        # 协作模式下, 数字人站在旁边，通过数字人按钮决定播放
                        message = json.dumps({
                            "tasks": "playlist",
                            "playlist": [
                                {"video": ppt_monitor.get_idle_video_file(ppt_page), "loop": -1}
                            ]
                        })
                        await handler.send_to_clients(message)
                        # 协作模式: 如果当前数字人状态为 playing, 则不做处理
                        
                        pass
                        # TODO:
                    elif cfg.config["work_mode"] == "auto":
                        # 自动模式
                        # 开始播放视频
                        logger.info("检测到自动模式, 数字人状态: %s", avatar_status)
                        if not avatar_status == "playing":
                            # 切换页面
                            ppt_monitor.goto_next_page()
                        # TODO:

                # 处理 avatar_command 变化
                if previous_config["avatar_command"] != cfg.config["avatar_command"]:
                    # 处理数字人指令
                    # 如果当前数字人是播放状态, 则发送暂停指令
                    if avatar_status == "playing":
                        logger.info("发送暂停指令")
                        message = json.dumps({
                            "tasks": "pause"
                        })
                        await handler.send_to_clients(message)
                        avatar_status = "pause"
                    elif avatar_status == "pause":
                        logger.info("发送恢复指令")
                        message = json.dumps({
                            "tasks": "play"
                        })
                        await handler.send_to_clients(message)
                        avatar_status = "playing"
                    else:
                        # 如果是其他状态,应该播放当前的页面
                        current_ppt_status = ppt_monitor.get_current_ppt_status()
                        if current_ppt_status["slides_count"] > 0:
                            ppt_page = current_ppt_status["present_slide_index"]
                            if ppt_page == -1:
                                ppt_page = current_ppt_status["edit_slide_index"]
                            message = json.dumps({
                                "tasks": "playlist",
                                "playlist": [
                                    {"video": ppt_monitor.get_slide_video_file(ppt_page), "loop": 1},
                                    {"video": ppt_monitor.get_idle_video_file(ppt_page), "loop": -1}
                                ]
                            })
                            await handler.send_to_clients(message)
                    # logger.info("发送任务")
                    # message = json.dumps({
                    #         "tasks": cfg.config["avatar_command"]
                    #     })
                    # await handler.send_to_clients(message)

            # 更新 previous_config
            previous_config = cfg.config.copy()
 
            # 监测 PPT 状态变化
            # 优先处理 PPT 激活文件的变化
            # ppt_page 表示当前的 ppt 页面编号
            ppt_changed = False
            ppt_page = -1

            current_ppt_status = ppt_monitor.get_current_ppt_status()
            if current_ppt_status["present_name"] != "" and current_ppt_status["present_name"] != previous_ppt_status["present_name"]:
                logger.info("检测到PPT发生变化,重新加载视频配置")
                ppt_monitor.update_slide_video_list(current_ppt_status["present_name"])
                ppt_changed = True
                if current_ppt_status["present_slide_index"] > 0:
                    ppt_page = current_ppt_status["present_slide_index"]
                else:
                    ppt_page = current_ppt_status["edit_slide_index"]
            else:
                if current_ppt_status["present_slide_index"] > 0:
                    if previous_ppt_status["present_slide_index"] != current_ppt_status["present_slide_index"]:
                        ppt_changed = True
                        ppt_page = current_ppt_status["present_slide_index"]
                else:
                    # 有一种情况: 退出播放时, 播放页面为-1，但是编辑页面不变，此时不应为变化
                    if previous_ppt_status["edit_slide_index"] != current_ppt_status["edit_slide_index"]:
                        ppt_changed = True
                        ppt_page = current_ppt_status["edit_slide_index"]

            # 如果发生变化
            if ppt_changed:
                logger.info("pptChanged: %s", current_ppt_status)
            
            # print(ppt_page)
            # 如果是自动模式,播放当前页面
            if cfg.config["work_mode"] == "auto":
                if ppt_page != -1:
                    message = json.dumps({
                        "tasks": "playlist",
                        "playlist": [
                            {"video": ppt_monitor.get_slide_video_file(ppt_page), "loop": 1},
                            {"video": ppt_monitor.get_idle_video_file(ppt_page), "loop": -1}
                        ]
                    })
                    await handler.send_to_clients(message)
            await asyncio.sleep(0.5)

            # 更新参数
            previous_ppt_status = current_ppt_status.copy()

def update_avatar_event(event):
    with open("config.json", "r", encoding="utf-8") as f:
        j = json.load(f)
    if not "avatar_event" in j:
        return
    j["avatar_event"] = event
    with open("config.json", "w", encoding="utf-8") as f:
        json.dump(j, f)

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
            logger.info("Client connected: %s", client_addr)
            try:
                async for message in websocket:
                    try:
                        logger.info("Received from %s: %s", client_addr, message)
                    except Exception as e:
                        logger.exception("Error printing received message: %s", e)
                    # 尝试解析为 JSON 并打印解析后的内容
                    try:
                        parsed = json.loads(message)
                        update_avatar_event(parsed)
                        logger.info("Parsed message: %s", parsed)
                    except Exception:
                        logger.exception("Error parsing received message: %s", e)
            except websockets.exceptions.ConnectionClosed:
                # 连接被客户端正常或异常关闭
                logger.info("Client %s closed", client_addr)
                pass
        finally:
            cls.clients.remove(websocket)
            logger.info("Client disconnected: %s", client_addr)
    
    @classmethod
    async def send_to_clients(cls, message):
        logger.info("Broadcasting message to %s clients: %s", len(cls.clients), message)
        if cls.clients:
            # 原版
            # await asyncio.wait([client.send(message) for client in cls.clients])
            # 某些版本下会出现
            # Passing coroutines coroutines is forbidden use tasks explicitly
            # 的提示, 需要先转换为任务
            tasks = [asyncio.create_task(client.send(message)) for client in cls.clients]
            await asyncio.wait(tasks)


if __name__ == "__main__":
    import tendo.singleton
    try:
        single = tendo.singleton.SingleInstance()
    except:
        logger.error("Another instance of the program is already running.")
        exit(1)
    asyncio.run(broadcast_slide_change())
