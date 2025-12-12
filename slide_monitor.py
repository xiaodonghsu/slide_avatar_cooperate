import win32com.client
import os
import json
import time

class SlideMonitor():
    def __init__(self, assets_base_dir = None):
        # 胶片放映程序
        self.__slide_app_list = ["PowerPoint.Application", "Kwpp.Application"]
        self.__slide_app = None
        self.slide_app_name = None
        # 场景加载管理: 周期性检查胶片配置是否更新, 及时加载胶片
        # 为避免端侧频繁修改胶片的配置, 加载配置文件设置最小修改间隔
        self.__scene_config_file = "scene.json"
        self.__scene_config_last_modified_time = None
        self.__previous_scene_config = None
        self.__scene_config = None
        self.scene_update_flag = False
        self.__minimal_modify_interval = 5
        self.__init_scene()
        # 资源管理
        # 启动方式 "start_new" "use_existing"
        self.slide_app_startup_method = None
        self.slide_show_active = False
        # self.__assets_base = self.load_assets()["assets_base"]
        self.__slide_video_config = "slide_video.json"
        self.__slide_index_prefix = "slide-"
        self.__idle_video_prefix = "idle"
        if assets_base_dir is None:
            assets_base_dir = self.__scene_config["assets_base"]
        if not assets_base_dir is None:
            assets_base_dir = os.path.abspath(assets_base_dir)
            if not os.path.isdir(assets_base_dir):
                raise Exception("资源路径不是有效路径:", assets_base_dir)
        self.__assets_base_dir = assets_base_dir
        # 前一个状态记录
        self.previous_presentation_name = None
        self.previous_edit_slide_index = None
        self.previous_show_slide_index = None
        # 当前状态记录
        self.current_presentation_name = None
        self.current_edit_slide_index = None
        self.current_show_slide_index = None
        # 只让 ppt_add 没有连接的信息出现一次
        self.__ppt_app_warning_flag = True

    def __init_scene(self):
        if self.__scene_config is None:
            with open(self.__scene_config_file, "r", encoding="utf-8") as f:
                self.__scene_config = json.load(f)
            self.__scene_config_last_modified_time = os.path.getmtime(self.__scene_config_file)
        if self.__previous_scene_config is None:
            self.__previous_scene_config = self.__scene_config.copy()

    def fresh_scene(self):
        self.scene_update_flag = False
        scene_config_modified_time = os.path.getmtime(self.__scene_config_file)
        if self.__scene_config_last_modified_time != scene_config_modified_time:
            print("检测到 scene 配置修改:", time.time() - scene_config_modified_time, "秒")
            if time.time() - scene_config_modified_time > self.__minimal_modify_interval:
                with open(self.__scene_config_file, "r", encoding="utf-8") as f:
                    self.__scene_config = json.load(f)
                print(self.__previous_scene_config["scene_active"], "->", self.__scene_config["scene_active"])
                if self.__scene_config["scene_active"] != self.__previous_scene_config["scene_active"]:
                    self.scene_update_flag = True
                print("更新文件修改时间及记忆配置")
                self.__previous_scene_config = self.__scene_config.copy()
                self.__scene_config_last_modified_time = scene_config_modified_time

    def get_active_asset_file(self):
        for asset in self.__scene_config["scene_list"]:
            if asset["name"] == self.__scene_config["scene_active"]:
                return os.path.join(os.path.split(os.path.abspath(__file__))[0], self.__scene_config["assets_base"], asset["file"])
        return None

    def get_active_scene_name(self):
        return self.__scene_config["scene_active"]

    def connect_slide_app(self, open_app=False):
        for app_name in self.__slide_app_list:
            try:
                self.slide_app = win32com.client.GetActiveObject(app_name)
                if self.slide_app:
                    self.slide_app_name = app_name
                    self.slide_app_startup_method = "use_existing"
                    break
            except Exception as e:
                self.slide_app = None

        if open_app and self.slide_app is None:
            for app_name in self.__slide_app_list:
                try:
                    self.slide_app = win32com.client.DispatchEx(app_name)
                    self.slide_app.DisplayAlerts = False
                    if self.slide_app:
                        self.slide_app_name = app_name
                        self.slide_app_startup_method = "start_new"
                        break
                except Exception as e:
                    self.slide_app = None

        if self.slide_app is None:
            if self.__ppt_app_warning_flag:
                self.__ppt_app_warning_flag = False
            return
        else:
             self.__ppt_app_warning_flag = True
        
        try:
            if not self.slide_app.Visible:
                self.slide_app.Visible = 1
        except Exception as e:
            self.slide_app = None
            print("Error in set visible:", e)

    def get_presentation_name(self):
        '''
        获取当前活动的演示文稿名称
        '''
        slide_app = self.get_slide_app()
        try:
            return slide_app.ActivePresentation.Name
        except:
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
        slide_app = self.get_slide_app()
        if slide_app:
            try:
                return slide_app.Presentations.Count
            except Exception as e:
                pass
        return -1

    def get_slides_count(self):
        if self.get_presentations_count()>0:
            try:
                presentation = self.slide_app.ActivePresentation
                return presentation.Slides.Count
            except:
                return -1
        return -1

    def get_edit_slide_index(self):
        '''
        获取编辑的幻灯片编号
        '''
        slide_app = self.get_slide_app()
        if slide_app:
            if slide_app.Presentations.Count > 0:
                presentation = self.slide_app.ActivePresentation
                if presentation:
                    try:
                        slide = presentation.Windows(1).View.Slide
                        return slide.SlideIndex if slide.SlideIndex is not None else -1
                    except:
                        pass
        return -1

    def get_show_slide_index(self):
        '''
        获取编辑的幻灯片编号
        '''
        slide_app = self.get_slide_app()
        if slide_app:
            if slide_app.Presentations.Count > 0:
                if slide_app.SlideShowWindows.Count > 0:
                    presentation = self.slide_app.ActivePresentation
                    if presentation:
                        try:
                            slide = presentation.SlideShowWindow.View.Slide
                            return slide.SlideIndex if slide.SlideIndex is not None else -1
                        except:
                            pass
        return -1

    def get_present_slide_index(self):
        '''
        获取播放的幻灯片编号
        '''
        if self.get_presentations_count()>0:
            try:
                presentation = self.slide_app.ActivePresentation
                slide = presentation.SlideShowWindow.View.Slide
                return slide.SlideIndex
            except:
                return -1

    def get_slide_video_list(self):
        '''
        获取当前幻灯片中的视频列表
        '''
        presentation_name = self.get_presentation_name()
        if presentation_name is None:
            return []

        # 读取演示文稿对应的视频列表配置
        assets_file = os.path.join(self.__assets_base_dir, self.__slide_video_config)
        with open(assets_file, "r", encoding='utf-8') as f:
            slide_video_config = json.load(f)

        # 搜索与当前演示文稿同名的视频列表
        for item in slide_video_config["slide_videos"]:
            if "name" in item and "videos" in item:
                if item["name"].endswith(presentation_name):
                    video_kv_list = item["videos"]
                    # 确认视频是否存在,并补充完整路径
                    for video_index in video_kv_list:
                        video_file = os.path.join(self.__assets_base_dir, video_kv_list[video_index])
                        if not os.path.exists(video_file):
                            video_kv_list[video_index] = ""
                        else:
                            video_kv_list[video_index] = os.path.abspath(video_file)
                    return video_kv_list

    def get_slide_video_file(self, slide_index):
        index = self.__slide_index_prefix + str(slide_index)
        return self.get_video_file(index)

    def get_idle_video_file(self, slide_index = None):
        index = self.__idle_video_prefix
        return self.get_video_file(index)

    def get_video_file(self, file_index):
        video_kv_list = self.get_slide_video_list()
        if not file_index in video_kv_list:
            return []
        return video_kv_list[file_index]

    def get_slide_app(self, open_app=False):
        if self.__slide_app is None:
            self.connect_slide_app(open_app)
        return self.slide_app

    def get_slide_show_index(self):
        slide_app = self.get_slide_app()
        if slide_app is None:
            return -1
        return slide_app.ActivePresentation.SlideShowWindow.View.Slide.SlideIndex

    def get_slide_edit_index(self):
        slide_app = self.get_slide_app()
        if slide_app is None:
            return -1
        return slide_app.ActivePresentation.Windows(1).View.Slide.SlideIndex

    def goto_page(self, dest_slide_index = 0):
        '''
        ppt 跳转到指定页面, -1, 指代上一页; 0 指代下一页; 正数指定特定页面
        如果是播放状态, 跳转播放
        如果是非播放状态, 则跳转编辑的页面
        '''
        if self.get_presentations_count() > 0:
            current_show_slide_index = self.get_show_slide_index()
            current_edit_slide_index = self.get_edit_slide_index()
            current_slides_count = self.get_slides_count()
            if current_show_slide_index > 0:
                # 下一页

                if dest_slide_index == 0:
                    if current_show_slide_index < current_slides_count:
                        self.slide_app.ActivePresentation.SlideShowWindow.View.Next()
                elif dest_slide_index == -1:
                    # 上一页
                    if current_show_slide_index > 1:
                        self.slide_app.ActivePresentation.SlideShowWindow.View.Previous()
                else:
                    if dest_slide_index >= 1 and dest_slide_index <= current_slides_count:
                        self.slide_app.ActivePresentation.SlideShowWindow.View.GotoSlide(dest_slide_index)
            elif current_edit_slide_index > 0:
                    # 下一页
                    if dest_slide_index == 0:
                        if current_edit_slide_index < current_slides_count:
                            self.slide_app.ActivePresentation.Slides(current_edit_slide_index + 1).Select()
                    elif dest_slide_index == -1:
                        # 上一页
                        if current_edit_slide_index > 1:
                            self.slide_app.ActivePresentation.Slides(current_edit_slide_index - 1).Select()
                    else:
                        if dest_slide_index >= 1 and dest_slide_index <= current_slides_count :
                            self.slide_app.ActivePresentation.Slides(dest_slide_index).Select()

    def goto_next_page(self):
        self.goto_page(0)

    def goto_previous_page(self):
        self.goto_page(-1)

    def start_slideshow(self):
        if self.get_presentations_count() <= 0:
            return
        try:
            self.slide_app.ActivePresentation.SlideShowSettings.Run()
            self.slide_show_active = True
        except:
            pass

    def update_state(self):
        # 记录保存到previous中
        self.previous_presentation_name = self.current_presentation_name 
        self.previous_slide_edit_index = self.current_slide_edit_index
        self.previous_slide_show_index = self.current_slide_show_index
        # 记录新的状态
        self.current_presentation_name = self.get_presentation_name()
        self.current_slide_edit_index = self.get_slide_edit_index()
        self.current_slide_show_index = self.get_slide_show_index()

    # 启动应用,加载加载胶片
    def open_presentation(self):
        slide_app = self.get_slide_app()
        if slide_app is None:
            slide_app = self.get_slide_app(open_app=True)
        if slide_app is None:
            return

        active_presentation_file = self.get_active_asset_file()
        print("需要放映的文档:", active_presentation_file)
        if active_presentation_file is None:
            print("尚未配置活跃演示文稿")
            return

        # 关闭所有的文档
        for presentation in slide_app.Presentations:
            presentation.Close()

        # active_presentation_name = os.path.split(active_presentation_file)[-1]
        # # 取得当前打开的文档名称的列表
        # presentations_name = []
        # for presentation in slide_app.Presentations:
        #     presentations_name.append(presentation.Name)
        # print("已打开的文档列表: ", presentations_name)
        # print("活跃演示文稿: ", active_presentation_name)
        # # 关闭与目标不一致的文档
        # for item in presentations_name:
        #     # 比较名称，不匹配则关闭
        #     print("比较: ", item, active_presentation_name, item != active_presentation_name)
        #     if item != active_presentation_name:
        #         print("关闭文档:", item)
        #         slide_app.Presentations[item].Close()
        
        print("尝试打开文档:", active_presentation_file)
        try:
            slide_app.Presentations.Open(active_presentation_file)
        except Exception as e:
            print("打开文档失败:", e)

    def start_slide_show(self):
        count_down = 120
        while True:
            print(count_down)
            if self.get_slide_app(True) is None:
                print("获取演示应用失败")
                break
            if self.get_presentations_count() < 0:
                print("尚未检测到演示播放程序")
            if self.get_presentations_count() >= 0:
                print("检测到演示播放程序, 尝试加载文档")
                self.open_presentation()
                print("演示文档已加载")
                if self.slide_app.SlideShowWindows.Count == 0:
                    print("演示文档开始放映")
                    self.slide_app.ActivePresentation.SlideShowSettings.Run()
                print("等待演示文档放映")
                time.sleep(3)
                if self.slide_app.SlideShowWindows.Count > 0:
                    print("已开始放映, 激活放映窗口")
                    self.slide_app.ActivePresentation.SlideShowWindow.Activate()
                    break
            time.sleep(1)
            count_down -= 1
            if count_down == 0:
                print("无PPT打开")
                break
