
import asyncio
import json
import os
import websockets
from config_monitor import Config
from slide_monitor import SlideMonitor

import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path

# Logger setup: write logs to project root `monitor.log`
if not os.path.exists("log"):
    os.mkdir("log")
LOG_FILE = os.path.join("log", "monitor.log")
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

global slide_monitor
slide_monitor = None

def parse_event(event: dict) -> str:
    '''
    根据事件, 返回数字人的状态
    idle
    playing
    unknown
    '''
    global slide_monitor
    if "event" in event and "type" in event and "src" in event:
        if event["event"] == "started":
            if event["type"] == "video":
                if event["src"] == slide_monitor.get_idle_video_file():
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
    cfg.fresh()
    previous_config = cfg.config.copy()
    logger.info("配置文件: %s", str(cfg.config))
    logger.info("初始化 Presentation 监控...")
    global slide_monitor
    slide_monitor = SlideMonitor()
    # slide_monitor.connect_slide_app()
    # presents_count = slide_monitor.get_presentations_count()

    # 启动应用
    # slide_monitor.open_presentation()
    
    message = json.dumps({
            "tasks": "text",
            "text": f"启动场景{slide_monitor.get_active_scene_name()}",
            "duration": 2
        })
    await handler.send_to_clients(message)
    
    slide_monitor.start_slide_show()
    
    # 监测初始化状态
    previous_present_name = slide_monitor.get_presentation_name()
    previous_edit_slide_index = slide_monitor.get_edit_slide_index()
    previous_show_slide_index = slide_monitor.get_show_slide_index()
    logger.info("PPT应用名称: %s, PPT文件名: %s, 编辑页面: %i, 放映页面: %i", 
                slide_monitor.slide_app_name,
                previous_present_name,
                previous_edit_slide_index,
                previous_show_slide_index)

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
            # 优先处理事件,减少事件问题导致状态的变化
            slide_monitor.fresh_scene()
            if slide_monitor.scene_update_flag:
                logger.info("检测到场景更新, 加载相关配置...")
                message = json.dumps({
                        "tasks": "text",
                        "text": f"切换到场景{slide_monitor.get_active_scene_name()}",
                        "duration": 2
                    })
                await handler.send_to_clients(message)
                
                # slide_monitor.open_presentation()
                slide_monitor.start_slide_show()

            cfg.fresh()
            if cfg.isFresh:
                if previous_config["avatar_event"] != cfg.config["avatar_event"]:
                    logger.info("数字人事件: %s", cfg.config["avatar_event"])
                    # 根据工作模式处理事件
                    avatar_status = parse_event(cfg.config["avatar_event"])
                    logger.info("数字人状态切换: %s", avatar_status)
                    if avatar_status == "idle":
                        if cfg.config["work_mode"] == "auto":
                            slide_monitor.goto_next_page()
                    elif avatar_status == "playing":
                        pass
                    elif avatar_status == "unknown":
                        pass

                # 处理 work_mode 变化
                if cfg.config["work_mode_response"]["result"] != "success":
                    logger.info("Switching to %s mode.", cfg.config["work_mode"])
                    message = json.dumps({
                        "tasks": "text",
                        "text": cfg.config["work_mode"],
                        "duration": 2
                    })
                    await handler.send_to_clients(message)
                    if cfg.config["work_mode"] == "manual":
                        # 讲解员模式: 主要的讲解任务在讲解员。数字人不参与
                        # 发送空的播放列表以停止播放
                        message = json.dumps({
                            "tasks": "playlist",
                            "playlist": []
                        })
                        await handler.send_to_clients(message)
                        cfg.update_work_mode_response()
                    elif cfg.config["work_mode"] == "collaboration":
                        # 协作模式下, 数字人站在旁边，通过数字人按钮决定播放
                        message = json.dumps({
                            "tasks": "playlist",
                            "playlist": [
                                {"video": slide_monitor.get_idle_video_file(slide_page), "loop": -1}
                            ]
                        })
                        await handler.send_to_clients(message)
                        # 协作模式: 如果当前数字人状态为 playing, 则不做处理
                        cfg.update_work_mode_response()
                        pass
                        # TODO:
                    elif cfg.config["work_mode"] == "auto":
                        # 自动模式
                        # 开始播放视频
                        logger.info("检测到自动模式, 数字人状态: %s", avatar_status)
                        if not avatar_status == "playing":
                            # 切换页面
                            slide_monitor.goto_next_page()
                        # TODO:
                        cfg.update_work_mode_response()

                # 处理 avatar_command 变化
                if cfg.config["avatar_command_response"]["result"] != "success":
                    logger.info("Avatar command: %s mode.", cfg.config["avatar_command"])
                    # 处理数字人指令
                    # 如果当前数字人是播放状态, 则发送暂停指令
                    if cfg.config["avatar_command"]["command"] == "play/pause":
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
                            logger.info("数字人状态: %s, 播放当前页面", avatar_status)
                            # 如果是其他状态,应该播放当前的页面
                            slides_count = slide_monitor.get_slides_count()
                            if slides_count > 0:
                                slide_page = slide_monitor.get_show_slide_index()
                                if slide_page == -1:
                                    slide_page = slide_monitor.get_edit_slide_index()
                                message = json.dumps({
                                    "tasks": "playlist",
                                    "playlist": [
                                        {"video": slide_monitor.get_slide_video_file(slide_page), "loop": 1},
                                        {"video": slide_monitor.get_idle_video_file(slide_page), "loop": -1}
                                    ]
                                })
                                await handler.send_to_clients(message)
                    elif cfg.config["avatar_command"]["command"] == "stop":
                        # 停止命令，直接进入空闲视频状态
                        message = json.dumps({
                            "tasks": "playlist",
                            "playlist": [
                                {"video": slide_monitor.get_idle_video_file(slide_page), "loop": -1}
                            ]
                        })
                        await handler.send_to_clients(message)
                        avatar_status = "idle"
                    elif cfg.config["avatar_command"]["command"] == "text":
                        # 发送文本消息
                        message = json.dumps({
                            "tasks": "text",
                            "text": cfg.config["avatar_command"]["text"],
                            "duration": 5
                        })
                        await handler.send_to_clients(message)
                    else:
                        # 如果是其他状态,应该播放当前的页面
                        slides_count = slide_monitor.get_slides_count()
                        if slides_count > 0:
                            slide_page = slide_monitor.get_show_slide_index()
                            if slide_page == -1:
                                slide_page = slide_monitor.get_edit_slide_index()
                            message = json.dumps({
                                "tasks": "playlist",
                                "playlist": [
                                    {"video": slide_monitor.get_slide_video_file(slide_page), "loop": 1},
                                    {"video": slide_monitor.get_idle_video_file(slide_page), "loop": -1}
                                ]
                            })
                            await handler.send_to_clients(message)
                    cfg.update_avatar_command_response()
                    # logger.info("发送任务")
                    # message = json.dumps({
                    #         "tasks": cfg.config["avatar_command"]
                    #     })
                    # await handler.send_to_clients(message)

            # 更新 previous_config
            previous_config = cfg.config.copy()
 
            # 监测 PPT 状态变化
            # 优先处理 PPT 激活文件的变化
            # slide_page 表示当前的 ppt 页面编号
            slide_changed = False
            slide_page = -1

            current_present_name = slide_monitor.get_presentation_name()
            current_show_slide_index = slide_monitor.get_show_slide_index()
            current_edit_slide_index = slide_monitor.get_edit_slide_index()
            # logger.info("current - name: %s, show: %s, edit: %s", current_present_name, current_show_slide_index, current_edit_slide_index)
            if current_present_name != "" and current_present_name != previous_present_name:
                # logger.info("检测到当前 演示文件 发生变化")
                # slide_monitor.update_slide_video_list(current_present_name)
                slide_changed = True
                if current_show_slide_index > 0:
                    slide_page = current_show_slide_index
                else:
                    slide_page = current_edit_slide_index
            else:
                if current_show_slide_index > 0:
                    if current_show_slide_index != previous_show_slide_index:
                        slide_changed = True
                        slide_page = current_show_slide_index
                else:
                    # 有一种情况: 退出播放时, 播放页面为-1，但是编辑页面不变，此时不应为变化
                    if current_show_slide_index != previous_show_slide_index:
                        logger.info("退出播放状态")
                    if current_edit_slide_index != previous_edit_slide_index:
                        slide_changed = True
                        slide_page = current_edit_slide_index

            # 如果发生变化
            if slide_changed:
                logger.info("pptChanged - name: %s, show: %s, edit: %s", current_present_name, current_show_slide_index, current_edit_slide_index)
            
            # print(slide_page)
            # 如果是自动模式,播放当前页面
            if cfg.config["work_mode"] == "auto":
                if slide_page != -1:
                    message = json.dumps({
                        "tasks": "playlist",
                        "playlist": [
                            {"video": slide_monitor.get_slide_video_file(slide_page), "loop": 1},
                            {"video": slide_monitor.get_idle_video_file(slide_page), "loop": -1}
                        ]
                    })
                    await handler.send_to_clients(message)
            await asyncio.sleep(0.5)

            # 更新参数
            previous_present_name = current_present_name
            previous_show_slide_index = current_show_slide_index
            previous_edit_slide_index = current_edit_slide_index

def update_avatar_event(event):
    with open("config.json", "r", encoding="utf-8") as f:
        j = json.load(f)
    if not "avatar_event" in j:
        return
    j["avatar_event"] = event
    with open("config.json", "w", encoding="utf-8") as f:
        json.dump(j, f, ensure_ascii=False, indent=4)

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
        else:
            logger.warning("No clients connected")


if __name__ == "__main__":
    import tendo.singleton
    try:
        single = tendo.singleton.SingleInstance()
    except:
        logger.error("Another instance of the program is already running.")
        exit(1)
    # 调用 npm start 方式启动数字人任务
    asyncio.run(broadcast_slide_change())
