import unittest
import json
import os
import tempfile
import shutil
from unittest.mock import patch, MagicMock, mock_open

# 导入要测试的SlideMonitor类
from slide_monitor import SlideMonitor

class TestSlideMonitor(unittest.TestCase):
    """SlideMonitor类的单元测试"""
    
    def setUp(self):
        """在每个测试方法之前运行"""
        # 创建临时目录和测试文件
        self.temp_dir = tempfile.mkdtemp()
        
        # 创建scene.json配置文件
        self.scene_config = {
            "assets_base": "assets",
            "asset_active": "default_scene",
            "assets_list": [
                {
                    "name": "default_scene",
                    "file": "presentation.pptx"
                }
            ]
        }
        
        self.scene_file = os.path.join(self.temp_dir, "scene.json")
        with open(self.scene_file, 'w', encoding='utf-8') as f:
            json.dump(self.scene_config, f, ensure_ascii=False, indent=4)
        
        # 创建slide_video.json配置文件
        self.slide_video_config = {
            "slide_videos": [
                {
                    "name": "test_presentation.pptx",
                    "videos": {
                        "slide-1": "videos/slide1.webm",
                        "slide-2": "videos/slide2.webm",
                        "idle": "videos/idle.webm"
                    }
                }
            ]
        }
        
        # 创建资产目录结构
        self.assets_dir = os.path.join(self.temp_dir, "assets")
        os.makedirs(self.assets_dir, exist_ok=True)
        
        self.slide_video_file = os.path.join(self.assets_dir, "slide_video.json")
        with open(self.slide_video_file, 'w', encoding='utf-8') as f:
            json.dump(self.slide_video_config, f, ensure_ascii=False, indent=4)
        
        # 创建测试视频文件
        video_dir = os.path.join(self.assets_dir, "videos")
        os.makedirs(video_dir, exist_ok=True)
        
        for video_file in ["slide1.webm", "slide2.webm", "idle.webm"]:
            with open(os.path.join(video_dir, video_file), 'w') as f:
                f.write("fake video content")
        
        # 创建演示文稿文件
        presentation_file = os.path.join(self.assets_dir, "presentation.pptx")
        with open(presentation_file, 'w') as f:
            f.write("fake presentation content")

    def tearDown(self):
        """在每个测试方法之后运行"""
        # 清理临时目录
        shutil.rmtree(self.temp_dir)

    @patch('slide_monitor.win32com.client')
    def test_slide_monitor_initialization(self, mock_win32com):
        """测试SlideMonitor初始化"""
        monitor = SlideMonitor(assets_base_dir=self.assets_dir)
        
        self.assertEqual(monitor._SlideMonitor__assets_config_file, "scene.json")
        self.assertEqual(monitor._SlideMonitor__slide_video_config, "slide_video.json")
        self.assertEqual(monitor._SlideMonitor__assets_base_dir, self.assets_dir)
        self.assertIsNotNone(monitor._SlideMonitor__assets_config)
        self.assertIsNotNone(monitor._SlideMonitor__previous_assets_config)

    def test_init_assets(self):
        """测试初始化资产配置"""
        monitor = SlideMonitor(assets_base_dir=self.assets_dir)
        
        # 确保资产配置正确加载
        self.assertEqual(monitor._SlideMonitor__assets_config["asset_active"], "default_scene")
        self.assertEqual(monitor._SlideMonitor__assets_config["assets_base"], "assets")
        
        # 确保之前配置也正确设置
        self.assertEqual(monitor._SlideMonitor__previous_assets_config["asset_active"], "default_scene")

    @patch('os.path.getmtime')
    def test_fresh_assets_with_update(self, mock_getmtime):
        """测试检测资产配置更新"""
        monitor = SlideMonitor(assets_base_dir=self.assets_dir)
        
        # 模拟文件修改时间变化
        mock_getmtime.return_value = 1234567890  # 不同的修改时间
        
        monitor.fresh_assets()
        
        # 应该检测到更新
        self.assertTrue(monitor.assets_update_flag)

    @patch('os.path.getmtime')
    def test_fresh_assets_without_update(self, mock_getmtime):
        """测试资产配置无更新"""
        monitor = SlideMonitor(assets_base_dir=self.assets_dir)
        
        # 模拟相同的修改时间
        original_mtime = monitor._SlideMonitor__assets_config_last_modified_time
        mock_getmtime.return_value = original_mtime
        
        monitor.fresh_assets()
        
        # 不应该检测到更新
        self.assertFalse(monitor.assets_update_flag)

    def test_get_active_asset_file(self):
        """测试获取活动资产文件路径"""
        monitor = SlideMonitor(assets_base_dir=self.assets_dir)
        
        asset_file = monitor.get_active_asset_file()
        expected_path = os.path.join(self.assets_dir, "presentation.pptx")
        
        self.assertEqual(asset_file, os.path.abspath(expected_path))

    def test_get_active_asset_file_nonexistent_asset(self):
        """测试获取不存在的活动资产文件"""
        # 修改配置中的活动资产为不存在的
        invalid_config = self.scene_config.copy()
        invalid_config["asset_active"] = "nonexistent_asset"
        
        with patch('builtins.open', mock_open(read_data=json.dumps(invalid_config))):
            with patch('os.path.getmtime', return_value=1234567890):
                monitor = SlideMonitor(assets_base_dir=self.assets_dir)
                monitor._SlideMonitor__assets_config = invalid_config
                
                asset_file = monitor.get_active_asset_file()
                self.assertIsNone(asset_file)

    def test_get_slide_video_list(self):
        """测试获取幻灯片视频列表"""
        monitor = SlideMonitor(assets_base_dir=self.assets_dir)
        
        # 模拟演示文稿名称
        with patch.object(monitor, 'get_presentation_name', return_value="test_presentation.pptx"):
            video_list = monitor.get_slide_video_list()
            
            expected_list = {
                "slide-1": os.path.abspath(os.path.join(self.assets_dir, "videos/slide1.webm")),
                "slide-2": os.path.abspath(os.path.join(self.assets_dir, "videos/slide2.webm")),
                "idle": os.path.abspath(os.path.join(self.assets_dir, "videos/idle.webm"))
            }
            
            self.assertEqual(video_list, expected_list)

    def test_get_slide_video_list_nonexistent_presentation(self):
        """测试获取不存在的演示文稿的视频列表"""
        monitor = SlideMonitor(assets_base_dir=self.assets_dir)
        
        # 模拟不存在的演示文稿名称
        with patch.object(monitor, 'get_presentation_name', return_value="nonexistent.pptx"):
            video_list = monitor.get_slide_video_list()
            self.assertEqual(video_list, [])

    def test_get_slide_video_list_nonexistent_video_files(self):
        """测试获取包含不存在视频文件的视频列表"""
        # 修改配置中的视频文件路径为不存在的
        invalid_video_config = self.slide_video_config.copy()
        invalid_video_config["slide_videos"][0]["videos"]["slide-1"] = "videos/nonexistent.webm"
        
        with patch('builtins.open', mock_open(read_data=json.dumps(invalid_video_config))):
            with patch.object(monitor, 'get_presentation_name', return_value="test_presentation.pptx"):
                monitor = SlideMonitor(assets_base_dir=self.assets_dir)
                
                video_list = monitor.get_slide_video_list()
                self.assertEqual(video_list["slide-1"], "")  # 不存在的文件应该返回空字符串

    def test_get_slide_video_file(self):
        """测试获取特定幻灯片的视频文件"""
        monitor = SlideMonitor(assets_base_dir=self.assets_dir)
        
        # 模拟视频列表
        mock_video_list = {
            "slide-1": "/path/to/slide1.webm",
            "slide-2": "/path/to/slide2.webm"
        }
        
        with patch.object(monitor, 'get_slide_video_list', return_value=mock_video_list):
            video_file = monitor.get_slide_video_file(1)
            self.assertEqual(video_file, "/path/to/slide1.webm")

    def test_get_slide_video_file_nonexistent_slide(self):
        """测试获取不存在的幻灯片的视频文件"""
        monitor = SlideMonitor(assets_base_dir=self.assets_dir)
        
        # 模拟视频列表
        mock_video_list = {
            "slide-1": "/path/to/slide1.webm"
        }
        
        with patch.object(monitor, 'get_slide_video_list', return_value=mock_video_list):
            video_file = monitor.get_slide_video_file(2)  # 不存在的幻灯片
            self.assertEqual(video_file, [])  # 应该返回空列表

    def test_get_idle_video_file(self):
        """测试获取空闲视频文件"""
        monitor = SlideMonitor(assets_base_dir=self.assets_dir)
        
        # 模拟视频列表
        mock_video_list = {
            "idle": "/path/to/idle.webm"
        }
        
        with patch.object(monitor, 'get_slide_video_list', return_value=mock_video_list):
            video_file = monitor.get_idle_video_file()
            self.assertEqual(video_file, "/path/to/idle.webm")

    def test_get_video_file(self):
        """测试通用获取视频文件方法"""
        monitor = SlideMonitor(assets_base_dir=self.assets_dir)
        
        # 模拟视频列表
        mock_video_list = {
            "test-video": "/path/to/test.webm"
        }
        
        with patch.object(monitor, 'get_slide_video_list', return_value=mock_video_list):
            video_file = monitor.get_video_file("test-video")
            self.assertEqual(video_file, "/path/to/test.webm")

    @patch('slide_monitor.win32com.client')
    def test_connect_slide_app_existing(self, mock_win32com):
        """测试连接已存在的幻灯片应用程序"""
        # 模拟已存在的应用程序
        mock_app = MagicMock()
        mock_win32com.GetActiveObject.return_value = mock_app
        
        monitor = SlideMonitor(assets_base_dir=self.assets_dir)
        monitor.connect_slide_app()
        
        self.assertEqual(monitor.slide_app, mock_app)
        self.assertEqual(monitor.slide_app_name, "PowerPoint.Application")
        self.assertEqual(monitor.slide_app_startup_method, "use_existing")

    @patch('slide_monitor.win32com.client')
    def test_connect_slide_app_new_instance(self, mock_win32com):
        """测试启动新的幻灯片应用程序实例"""
        # 模拟没有已存在的应用程序，但可以启动新实例
        mock_win32com.GetActiveObject.side_effect = Exception("Not found")
        mock_app = MagicMock()
        mock_win32com.DispatchEx.return_value = mock_app
        
        monitor = SlideMonitor(assets_base_dir=self.assets_dir)
        monitor.connect_slide_app(open_app=True)
        
        self.assertEqual(monitor.slide_app, mock_app)
        self.assertEqual(monitor.slide_app_name, "PowerPoint.Application")
        self.assertEqual(monitor.slide_app_startup_method, "start_new")

    @patch('slide_monitor.win32com.client')
    def test_connect_slide_app_failure(self, mock_win32com):
        """测试连接幻灯片应用程序失败"""
        # 模拟所有连接方式都失败
        mock_win32com.GetActiveObject.side_effect = Exception("Not found")
        mock_win32com.DispatchEx.side_effect = Exception("Cannot create")
        
        monitor = SlideMonitor(assets_base_dir=self.assets_dir)
        monitor.connect_slide_app(open_app=True)
        
        self.assertIsNone(monitor.slide_app)
        self.assertIsNone(monitor.slide_app_name)
        self.assertIsNone(monitor.slide_app_startup_method)

    def test_is_connected_with_connection(self):
        """测试检查连接状态（已连接）"""
        monitor = SlideMonitor(assets_base_dir=self.assets_dir)
        
        with patch.object(monitor, 'get_presentations_count', return_value=1):
            self.assertTrue(monitor.isConnected())

    def test_is_connected_without_connection(self):
        """测试检查连接状态（未连接）"""
        monitor = SlideMonitor(assets_base_dir=self.assets_dir)
        
        with patch.object(monitor, 'get_presentations_count', return_value=0):
            self.assertFalse(monitor.isConnected())

if __name__ == '__main__':
    unittest.main()