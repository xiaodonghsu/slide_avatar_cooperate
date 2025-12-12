import unittest
import json
import asyncio
import websockets
from unittest.mock import patch, MagicMock, AsyncMock

# 导入要测试的模块
from monitor import handler, parse_event, update_avatar_event

class TestMonitorFunctions(unittest.TestCase):
    """monitor.py中函数的单元测试"""
    
    def setUp(self):
        """在每个测试方法之前运行"""
        # 清空handler的客户端集合
        handler.clients.clear()

    def test_parse_event_started_video_idle(self):
        """测试解析数字人开始播放空闲视频事件"""
        event = {
            "event": "started",
            "type": "video",
            "src": "/path/to/idle_video.webm"
        }
        
        # 模拟slide_monitor的get_idle_video_file方法
        mock_slide_monitor = MagicMock()
        mock_slide_monitor.get_idle_video_file.return_value = "/path/to/idle_video.webm"
        
        with patch('monitor.slide_monitor', mock_slide_monitor):
            result = parse_event(event)
            self.assertEqual(result, "idle")

    def test_parse_event_started_video_playing(self):
        """测试解析数字人开始播放普通视频事件"""
        event = {
            "event": "started",
            "type": "video",
            "src": "/path/to/other_video.webm"
        }
        
        # 模拟slide_monitor的get_idle_video_file方法
        mock_slide_monitor = MagicMock()
        mock_slide_monitor.get_idle_video_file.return_value = "/path/to/idle_video.webm"
        
        with patch('monitor.slide_monitor', mock_slide_monitor):
            result = parse_event(event)
            self.assertEqual(result, "playing")

    def test_parse_event_finished(self):
        """测试解析数字人播放完成事件"""
        event = {
            "event": "finished",
            "type": "video",
            "src": "/path/to/video.webm"
        }
        
        result = parse_event(event)
        self.assertEqual(result, "idle")

    def test_parse_event_unknown(self):
        """测试解析未知事件"""
        event = {
            "event": "unknown_event",
            "type": "unknown_type"
        }
        
        result = parse_event(event)
        self.assertEqual(result, "unknown")

    def test_parse_event_missing_fields(self):
        """测试解析缺少必要字段的事件"""
        event = {
            "event": "started"
            # 缺少type和src字段
        }
        
        result = parse_event(event)
        self.assertEqual(result, "unknown")

    @patch('builtins.open', new_callable=MagicMock)
    @patch('json.load')
    @patch('json.dump')
    def test_update_avatar_event(self, mock_json_dump, mock_json_load, mock_open):
        """测试更新头像事件配置"""
        # 模拟现有的配置
        existing_config = {
            "work_mode": "auto",
            "avatar_event": {"old": "event"}
        }
        mock_json_load.return_value = existing_config
        
        new_event = {"new": "event_data"}
        update_avatar_event(new_event)
        
        # 验证配置被正确更新
        expected_config = existing_config.copy()
        expected_config["avatar_event"] = new_event
        
        mock_json_dump.assert_called_once_with(expected_config, mock_open.return_value.__enter__.return_value, ensure_ascii=False, indent=4)

class TestHandlerClass(unittest.TestCase):
    """handler类的单元测试"""
    
    def setUp(self):
        """在每个测试方法之前运行"""
        # 清空handler的客户端集合
        handler.clients.clear()

    def test_handler_initialization(self):
        """测试handler类初始化"""
        self.assertEqual(handler.clients, set())

    @patch('monitor.logger')
    async def test_handler_adds_client_on_connection(self, mock_logger):
        """测试新客户端连接时被添加到集合中"""
        mock_websocket = AsyncMock()
        mock_websocket.remote_address = ('127.0.0.1', 12345)
        
        # 创建handler协程
        handler_coroutine = handler.handler(mock_websocket)
        
        # 运行协程直到第一个await
        try:
            handler_coroutine.send(None)
        except StopIteration:
            pass
        
        # 验证客户端被添加
        self.assertIn(mock_websocket, handler.clients)
        mock_logger.info.assert_called_once_with("Client connected: %s", ('127.0.0.1', 12345))

    @patch('monitor.logger')
    async def test_handler_removes_client_on_disconnect(self, mock_logger):
        """测试客户端断开连接时从集合中移除"""
        mock_websocket = AsyncMock()
        mock_websocket.remote_address = ('127.0.0.1', 12345)
        
        # 添加客户端
        handler.clients.add(mock_websocket)
        
        # 模拟连接关闭异常
        mock_websocket.__aiter__.return_value = AsyncMock()
        mock_websocket.__aiter__.return_value.__anext__.side_effect = websockets.exceptions.ConnectionClosed(1000, "normal")
        
        # 运行handler
        try:
            await handler.handler(mock_websocket)
        except:
            pass
        
        # 验证客户端被移除
        self.assertNotIn(mock_websocket, handler.clients)
        mock_logger.info.assert_any_call("Client closed: %s", ('127.0.0.1', 12345))

    @patch('monitor.logger')
    @patch('monitor.update_avatar_event')
    async def test_handler_processes_json_message(self, mock_update_avatar_event, mock_logger):
        """测试handler处理JSON消息"""
        mock_websocket = AsyncMock()
        mock_websocket.remote_address = ('127.0.0.1', 12345)
        
        # 模拟消息迭代器
        test_message = json.dumps({"event": "started", "type": "video"})
        mock_websocket.__aiter__.return_value = AsyncMock()
        mock_websocket.__aiter__.return_value.__anext__.return_value = test_message
        
        # 运行handler
        try:
            await handler.handler(mock_websocket)
        except StopIteration:
            pass
        
        # 验证消息被处理和记录
        mock_logger.info.assert_any_call("Received from %s: %s", ('127.0.0.1', 12345), test_message)
        mock_update_avatar_event.assert_called_once_with({"event": "started", "type": "video"})

    @patch('monitor.logger')
    async def test_handler_processes_invalid_json_message(self, mock_logger):
        """测试handler处理无效的JSON消息"""
        mock_websocket = AsyncMock()
        mock_websocket.remote_address = ('127.0.0.1', 12345)
        
        # 模拟无效的JSON消息
        invalid_message = "invalid json"
        mock_websocket.__aiter__.return_value = AsyncMock()
        mock_websocket.__aiter__.return_value.__anext__.return_value = invalid_message
        
        # 运行handler
        try:
            await handler.handler(mock_websocket)
        except StopIteration:
            pass
        
        # 验证错误被记录
        mock_logger.exception.assert_called()

    @patch('monitor.logger')
    async def test_send_to_clients_with_clients(self, mock_logger):
        """测试向所有客户端发送消息（有客户端连接）"""
        mock_websocket1 = AsyncMock()
        mock_websocket2 = AsyncMock()
        
        handler.clients.update([mock_websocket1, mock_websocket2])
        
        test_message = "test message"
        
        # 运行send_to_clients
        await handler.send_to_clients(test_message)
        
        # 验证消息发送给所有客户端
        mock_websocket1.send.assert_called_once_with(test_message)
        mock_websocket2.send.assert_called_once_with(test_message)
        mock_logger.info.assert_called_once_with("Broadcasting message to %s clients: %s", 2, test_message)

    @patch('monitor.logger')
    async def test_send_to_clients_no_clients(self, mock_logger):
        """测试向所有客户端发送消息（无客户端连接）"""
        test_message = "test message"
        
        # 运行send_to_clients
        await handler.send_to_clients(test_message)
        
        # 验证警告被记录
        mock_logger.warning.assert_called_once_with("No clients connected")

class TestAsyncFunctions(unittest.IsolatedAsyncioTestCase):
    """异步函数的单元测试"""
    
    @patch('monitor.handler.send_to_clients')
    @patch('monitor.Config')
    @patch('monitor.SlideMonitor')
    async def test_broadcast_slide_change_basic_flow(self, mock_slide_monitor_class, mock_config_class, mock_send_to_clients):
        """测试broadcast_slide_change的基本流程"""
        # 模拟配置对象
        mock_config = MagicMock()
        mock_config.config = {
            "server_host": "127.0.0.1",
            "websocket_port": 5678,
            "work_mode": "auto",
            "avatar_event": {},
            "avatar_command": {},
            "avatar_command_response": {"result": "pending"},
            "work_mode_response": {"result": "pending"}
        }
        mock_config.fresh.return_value = None
        mock_config.isFresh = False
        mock_config_class.return_value = mock_config
        
        # 模拟幻灯片监控器
        mock_slide_monitor = MagicMock()
        mock_slide_monitor.get_presentation_name.return_value = "test.pptx"
        mock_slide_monitor.get_edit_slide_index.return_value = 1
        mock_slide_monitor.get_show_slide_index.return_value = 1
        mock_slide_monitor.fresh_assets.return_value = None
        mock_slide_monitor_class.return_value = mock_slide_monitor
        
        # 模拟WebSocket服务器
        with patch('monitor.websockets.serve') as mock_serve:
            mock_server = AsyncMock()
            mock_serve.return_value = mock_server
            
            # 运行函数（短时间内）
            try:
                task = asyncio.create_task(monitor.broadcast_slide_change())
                await asyncio.sleep(0.1)  # 短暂运行
                task.cancel()
                await task
            except asyncio.CancelledError:
                pass
            
            # 验证WebSocket服务器启动
            mock_serve.assert_called_once()

if __name__ == '__main__':
    unittest.main()