import win32com.client
import asyncio
import json
import websockets
import time

WORK_MODE_FILE = "work-mode.json"
WS_PORT = 8765

def get_work_mode():
    '''
    从文件获取当前的工作模式，如果不存在则默认为手动模式
    '''
    try:
        with open(WORK_MODE_FILE, "r") as f:
            return json.load(f)
    except:
        with open(WORK_MODE_FILE, "w") as f:
            json.dump({"mode": "manual", "loop": False}, f)
        return {"mode": "manual", "loop": False}

class PowerPointMonitor():
    def __init__(self):
        self.ppt_app = None
        self.connect_powerpoint()
        self.slide_show_active = False
        self.presentation_name = None

    def connect_powerpoint(self):
        try:
            self.ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
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
    ppt_monitor = PowerPointMonitor()
    ppt_name = ppt_monitor.get_presentation_name()
    if not ppt_name  is None:
        print("Current PPT", ppt_name)
    else:
        print("No PPT opened")

    previous_edit_slide_index = -1
    previous_present_slide_index = None

    async with websockets.serve(handler.handler, "localhost", WS_PORT):
        print(f"WebSocket server started at ws://localhost:{WS_PORT}")
        while True:
            work_mode = get_work_mode()
            # print(time.strftime("%H:%M:%S"), f"{work_mode=}")
            if work_mode["mode"] == "manual":
                await asyncio.sleep(1)
                continue

            present_count, edit_index, present_index = ppt_monitor.get_current_status()
            # print(f"{present_count}, {edit_index}, {present_index}")
            if present_count > 0:
                new_slide_index = -1
                if (present_index != previous_present_slide_index):
                    new_slide_index = present_index
                else:
                    if (edit_index != previous_edit_slide_index):
                        new_slide_index = edit_index
                previous_present_slide_index = present_index
                previous_edit_slide_index = edit_index
                # print(new_slide_index)
                if new_slide_index != -1:
                    print(time.strftime("%H-%M-%S"), f"Status: {present_count}, Current edit slide: {edit_index}, Current Present slide: {present_index}")
                    message = json.dumps({
                        "tasks": "playlist",
                        "playlist": [
                            {"video": f"../assets/videos/video{new_slide_index}.webm", "loop": 1},
                            {"video": "../assets/videos/idle.webm", "loop": 999}
                        ]
                    })
                    await handler.send_to_clients(message)
            await asyncio.sleep(0.5)

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
            await asyncio.wait([client.send(message) for client in cls.clients])

if __name__ == "__main__":
    asyncio.run(broadcast_slide_change())
