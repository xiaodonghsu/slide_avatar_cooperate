from slide_monitor import SlideMonitor

if __name__ == "__main__":
    s = SlideMonitor()
    s.connect_slide_app()
    print(f"{s.slide_app=}")
    print(f"{s.get_edit_slide_index()=}")
    print(f"{s.slide_app_startup_method=}")
    print(f"{s.get_presentation_name()=}")
    print(f"{s.get_show_slide_index()=}")
    print(f"{s.get_edit_slide_index()=}")
    print(f"{s.get_slide_video_file(1)=}")