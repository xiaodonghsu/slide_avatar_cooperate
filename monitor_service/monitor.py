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
            with open(self.__CONFIG_FILE, "r") as f:
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
        self.__ppt_app_warning_flag = True
        self.connect_powerpoint()

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

    def get_current_ppt_status(self):
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
        return self.slide_video_list[self.__slide_index_prefix + str(prensentation_index)]

    def get_idle_video_file(self, prensentation_index):
        return self.slide_video_list[self.__idle_video_prefix]

async def broadcast_slide_change():

    # 首次启动使用当前的配置作为基础配置
    logger.info("读取配置...")
    cfg = Config()
    previous_config = cfg.config.copy()
    previous_edit_slide_index = -1
    previous_present_slide_index = -1

    logger.info("初始化 Presentation 监控...")
    ppt_monitor = PowerPointMonitor()
    ppt_name = ppt_monitor.get_presentation_name()
    present_count, previous_edit_slide_index, previous_present_slide_index = ppt_monitor.get_current_ppt_status()
    if ppt_name is None:
        logger.warning("No PPT opened")
    else:
        logger.info("PPT应用名称: %s", ppt_monitor.ppt_app_name)
        logger.info("当前PPT文件名: %s", ppt_name)
        logger.info("编辑页面: %s, 播放页面: %s, 总页面数: %s", previous_edit_slide_index, previous_present_slide_index, present_count)

    # 启动 WebSocket 服务器
    async with websockets.serve(handler.handler, cfg.config["server_host"], cfg.config["websocket_port"]):
        logger.info('WebSocket server started at ws://%s:%s', cfg.config["server_host"], cfg.config["websocket_port"])
        while True:
            # 监测配置文件更新
            # 如果文件更新, 需要分析具体的变化:
            # 从 auto 或 collaboration 切换到 manual, 需要发送停止播放的消息
            cfg.update()
            if cfg.isUpdated:
                if cfg.config["work_mode"] == "manual":
                    logger.info("Switched to manual mode.")
                    if previous_config["work_mode"] != cfg.config["work_mode"]:
                        logger.info("%s Switched to %s mode.", time.strftime("%H-%M-%S"), cfg.config["work_mode"])
                        # 发送空的播放列表以停止播放
                        message = json.dumps({
                            "tasks": "playlist",
                            "playlist": []
                        })
                        await handler.send_to_clients(message)
                        pass

                # 更新后的配置处理检查avatar_status变化
                if previous_config["avatar_status"] != cfg.config["avatar_status"]:
                    logger.info("发送任务")
                    message = json.dumps({
                            "tasks": cfg.config["avatar_status"]
                        })
                    await handler.send_to_clients(message)
            # 更新 previous_config
            previous_config = cfg.config.copy()
 
            # 监测 PPT 状态变化, ppt_page 表示当前的 ppt 号码
            ppt_changed = False
            ppt_page = -1
            present_count, edit_index, present_index = ppt_monitor.get_current_ppt_status()
            if present_index > 0:
                if previous_present_slide_index != present_index:
                    ppt_changed = True
                    ppt_page = present_index
            else:
                if previous_edit_slide_index != edit_index:
                    ppt_changed = True
                    ppt_page = edit_index

            if ppt_changed:
                logger.info("页面总数: %s, 编辑页面: %s, 播放页面: %s", present_count, edit_index, present_index)
            
            # print(ppt_page)
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
            #     # present_count, edit_index, present_index = ppt_monitor.get_current_ppt_status()
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
                        logger.debug("Parsed message: %s", parsed)
                    except Exception:
                        # 非 JSON 消息则忽略解析错误
                        pass
            except websockets.exceptions.ConnectionClosed:
                # 连接被客户端正常或异常关闭
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
    asyncio.run(broadcast_slide_change())
