document.addEventListener("DOMContentLoaded", () => {
    const video = document.getElementById("avatar");
    let currentVideo = null;
    let currentPage = null;

    // ---- 向 Python 发消息 ----
    function notifyPython(event) {
        try {
            if (window.pptWS && typeof window.pptWS.send === 'function') {
                window.pptWS.send(event);
                console.log("-> Python :", event);
            } else {
                console.warn('notifyPython: pptWS.send is not a available');
            }
        } catch (e) {
            console.error('notifyPython error', e);
        }
    }

    // ---- 播放函数 ----
    function play(videoPath, loop = false) {
        currentVideo = videoPath;
        video.src = videoPath;
        video.loop = loop;
        video.muted = false; // 取消静音
        video.volume = 1;   // 设置最大音量

        video.onended = () => {
            notifyPython({
                event: "finished",
                video: currentVideo,
                page: currentPage
            });
        };

        video.play().catch(err => console.error("play video error: ", err));
    }

    // ---- WebSocket 收到 Python 指令 ----
    // {"tasks": "play",
    //     "playlist": [
    //         {"video": "../assets/videos/video1.webm", "loop": 1, "left": 1200, "top": 700, "width": 100, "height": 300}, //播放1次
    //         {"image": "../assets/videos/video2.jpeg", "loop": 3}, //持续3秒
    //         {"video": "../assets/videos/idle.webm", "loop": 999}  //无限循环
    //          ...
    //          ]
    // }
    function handleCommand(data) {
        console.log("Python command -> " + data.toString('utf8'));

        if (data.command === "play") {
            play(data.uri, data.loop);
        }

        if (data.command === "idle") {
            currentPage = null;
            play(data.video, true);
        }

        if (data.command === "stop") {
            video.pause();
            video.src = "";
            currentVideo = null;
        }

        if (data.command === "pause") {
            video.pause();
            video.src = "";
            currentVideo = null;
        }
    }

    // ---- 连接 WebSocket ----
    window.pptWS.connect("ws://localhost:8765", (message) => {
        handleCommand(message);
    });

    // 启动时进入 idle
    setTimeout(() => {
        notifyPython({ event: "ready" });
    }, 500);
});
