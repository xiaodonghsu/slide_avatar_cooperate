from ntpath import abspath
import os
import json
from dotenv import load_dotenv

class AssetManager():
    def __init__(self):
        self.__slide_video_config_file = "slide_video.json"
        self.__scene_config_file = "scene.json"
        self.__slide_index_prefix = "slide-"
        self.__idle_video_prefix = "idle"

        load_dotenv()
        # 获取资源路径
        self.__assets_base_dir = os.getenv("ASSETS_BASE_DIR")
        # print(self.__assets_base_dir)
        # 读取角色配置
        self.__role = os.getenv("SCENE_ROLE")
        if self.__role is None:
            raise Exception("ROLE not found in environment variables")

        if not os.path.exists(self.__assets_base_dir):
            raise Exception(f"Assets base directory not found: {self.__assets_base_dir}")
        
        # 检查 slide_video.json 文件是否存在
        if not os.path.exists(os.path.join(self.__assets_base_dir, self.__slide_video_config_file)):
            raise Exception(f"Slide video config file not found: {self.__slide_video_config_file}")
        
        # 读取 slide_video.json 文件
        with open(os.path.join(self.__assets_base_dir, self.__slide_video_config_file), "r", encoding="utf-8") as f:
            self.__slide_video_config = json.load(f)

        # 检查 scene.json 文件是否存在
        # scene.json 格式
        '''
        {
            "scene_list": 
            [
                {
                "scene_name": "",
                "roles": [
                    {"role": "",
                    "script": ""
                    }]
                }
            ]
        }
        '''
        if not os.path.exists(os.path.join(self.__assets_base_dir, self.__scene_config_file)):
            raise Exception(f"Scene config file not found: {self.__scene_config_file}")

        # 读取 scene.json 文件
        with open(os.path.join(self.__assets_base_dir, self.__scene_config_file), "r", encoding="utf-8") as f:
            self.__scene_config = json.load(f)

  
    def get_slide_file(self, scene_name):
        '''
        获取当前幻灯片中的视频列表
        '''
        for scene in self.__scene_config["scene_list"]:
            if scene["scene_name"] == scene_name:
                for role in scene["roles"]:
                    if role["role"] == self.__role:
                        script_file = os.path.abspath(os.path.join(self.__assets_base_dir, role['script']))
                        if os.path.exists(script_file):
                            return script_file
        return None

    def get_slide_video_file(self, slide_name, slide_index=None):

        '''

        获取指定幻灯片的视频文件

        '''
        for sv in self.__slide_video_config["slide_videos"]:
            if sv["name"] == slide_name:
                file_index = None
                if slide_index is None:
                    file_index = self.__idle_video_prefix
                else:
                    file_index = self.__slide_index_prefix + str(slide_index)
                if file_index in sv["videos"]:
                    file_name = os.path.join(self.__assets_base_dir, sv["videos"][file_index])
                    if os.path.exists(file_name):
                        file_name = os.path.abspath(file_name)
                        return file_name
        return None


    def get_scenes(self):
        '''
        获取所有场景列表
        '''
        return [scene["scene_name"] for scene in self.__scene_config["scene_list"]]

    def get_next_scene(self, scene_name):

        '''
        获取下一个场景
        '''
        try:
            current_index = [scene["scene_name"] for scene in self.__scene_config["scene_list"]].index(scene_name)

            next_index = (current_index + 1) % len(self.__scene_config["scene_list"])

            return self.__scene_config["scene_list"][next_index]["scene_name"]
        except ValueError:
            return None