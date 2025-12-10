import win32com.client
import os
import json

class SlideMonitor():
    def __init__(self, assets_base_dir = None):
        self.__slide_app_list = ["PowerPoint.Application", "Kwpp.Application"]
        self.__slide_app = None
        self.slide_app_name = None
        # 启动方式 "start_new" "use_existing"
        self.slide_app_startup_method = None
        self.slide_show_active = False
        self.__assets_base_path_file = "assets_base_path.txt"
        self.__slide_video_config = "slide_video.json"
        self.__slide_index_prefix = "slide-"
        self.__idle_video_prefix = "idle"
        if assets_base_dir is None:
            current_path = os.path.split(os.path.abspath(__file__))[0]
            try:
                with open(os.path.join(current_path, self.__assets_base_path_file), "r") as f:
                    assets_base_dir = f.read().strip()
            except Exception as e:
                raise Exception("资源文件不存在:", e)
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
            for app_name in self._ppt_app_list:
                try:
                    self.ppt_app = win32com.client.DispatchEx(app_name)
                    self.ppt_app.DisplayAlerts = False
                    if self.ppt_app:
                        self.ppt_app_name = app_name
                        self.slide_app_startup_method = "start_new"
                        break
                except Exception as e:
                    self.ppt_app = None

        if self.slide_app is None:
            if self.__ppt_app_warning_flag:
                self.__ppt_app_warning_flag = False
            return
        else:
             self.__ppt_app_warning_flag = True
        
        if not self.slide_app.Visible:
            try:
                self.slide_app.Visible = 1
            except Exception as e:
                print(e)

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
            return slide_app.Presentations.Count
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
                slide = presentation.Windows(1).View.Slide
                return slide.SlideIndex if slide.SlideIndex is not None else -1
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
                    slide = presentation.SlideShowWindow.View.Slide
                    return slide.SlideIndex if slide.SlideIndex is not None else -1
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

    # def get_current_ppt_status(self):
    #     '''
    #     获取当前幻灯片的播放状态和编号, 返回 
    #     {"present_count": 打开的ppt数量,
    #     "present_name": ppt名字,
    #     "slides_count": 胶片数量,
    #     "edit_slide_index": 当前编辑的胶片的索引,
    #     "present_slide_index": 放映的胶片的索引}
    #     编号为 -1 无效
    #     '''
    #     # 检查当前连接状态
    #     current_ppt_status = {
    #         "present_count": -1,
    #         "present_name": "",
    #         "slides_count": -1,
    #         "edit_slide_index": -1,
    #         "present_slide_index": -1
    #     }
    #     # 如果没有连接 PowerPoint，则尝试连接
    #     if not self.isConnected():
    #         self.connect_slide_app()
    #     present_count = self.get_presentations_count()
    #     # 根据当前连接状态返回结果
    #     if present_count == -1:
    #         return current_ppt_status
    #     if present_count == 0:
    #         current_ppt_status["present_count"] = present_count
    #         return current_ppt_status
    #     if present_count > 0:
    #         current_ppt_status["present_count"] = present_count
    #         presentation = self.slide_app.ActivePresentation
    #         current_ppt_status["present_name"] = presentation.Name
    #         current_ppt_status["slides_count"] = presentation.Slides.Count
    #         current_ppt_status["edit_slide_index"] = presentation.Windows(1).View.Slide.SlideIndex
    #         # 则获取当前幻灯片编号
    #         present_slide_index = -1
    #         try:
    #             present_slide_index = presentation.SlideShowWindow.View.Slide.SlideIndex
    #             current_ppt_status["present_slide_index"] = present_slide_index
    #         except:
    #             pass
    #         return current_ppt_status

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

    def get_slide_app(self):
        if self.__slide_app is None:
            self.connect_slide_app()
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
