document.addEventListener("DOMContentLoaded", () => {
    const video = document.getElementById("avatar");
    const img = document.getElementById("avatarImg");
    let currentVideo = null;
    let currentPage = null;
    // playlist state
    let playlist = [];
    let currentIndex = -1;
    let playingItem = null; // the current playlist item

    // video loop remaining (for finite loops)
    let videoLoopRemaining = 0;

    // image timer state
    let imageTimer = null;
    let imageRemaining = 0; // ms
    let imageStartAt = 0; // timestamp when image timer started

    let isPaused = false;

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
    //         {"video": "../assets/videos/idle.webm", "loop": -1}  //无限循环
    //          ...
    //          ]
    // }
    // ----- Playlist / controller handling -----
    function clearPlaybackState() {
        // stop video
        try {
            video.pause();
            video.removeAttribute('src');
            video.load();
        } catch (e) {
            console.warn('clearPlaybackState video error', e);
        }

        // hide image and clear timers
        if (imageTimer) {
            clearTimeout(imageTimer);
            imageTimer = null;
        }
        img.style.display = 'none';

        playingItem = null;
        videoLoopRemaining = 0;
        imageRemaining = 0;
        imageStartAt = 0;
        currentVideo = null;
    }

    function stopAll() {
        playlist = [];
        currentIndex = -1;
        clearPlaybackState();
    }

    function advanceToNext() {
        if (!playlist || playlist.length === 0) {
            clearPlaybackState();
            return;
        }

        currentIndex += 1;
        if (currentIndex >= playlist.length) {
            // reached the end: stop
            clearPlaybackState();
            return;
        }

        startCurrentItem();
    }

    function startCurrentItem() {
        clearPlaybackState();
        if (currentIndex < 0 || currentIndex >= playlist.length) return;
        const item = playlist[currentIndex];
        playingItem = item;

        if (item.image) {
            // show image for loop seconds; loop==-1 means infinite, loop==0 means skip
            img.src = item.image;
            img.style.display = 'block';
            video.style.display = 'none';

            // loop semantics for image:
            // -1 => infinite display
            // 0  => skip this item
            // >0 => treat as seconds to display
            const loopVal = Number(item.loop);
            if (loopVal === -1) {
                imageRemaining = Infinity;
                imageTimer = null;
                notifyPython({ event: 'started', type: 'image', src: item.image, index: currentIndex });
            } else if (loopVal === 0) {
                // skip image immediately
                advanceToNext();
                return;
            } else {
                const durationMs = (loopVal > 0 ? loopVal : 1) * 1000;
                imageRemaining = durationMs;
                imageStartAt = Date.now();
                imageTimer = setTimeout(() => {
                    notifyPython({ event: 'finished', type: 'image', src: item.image, index: currentIndex });
                    imageTimer = null;
                    advanceToNext();
                }, imageRemaining);
                notifyPython({ event: 'started', type: 'image', src: item.image, index: currentIndex });
            }

        } else if (item.video) {
            img.style.display = 'none';
            video.style.display = 'block';
            currentVideo = item.video;

            // prepare loop behavior for video:
            // -1 => infinite loop
            // 0  => skip this item
            // >0 => loop that many times
            const loopVal = Number(item.loop);
            if (loopVal === -1) {
                video.loop = true;
                videoLoopRemaining = Infinity;
            } else if (loopVal === 0) {
                // skip this video
                advanceToNext();
                return;
            } else {
                video.loop = false;
                videoLoopRemaining = Math.max(1, loopVal || 1);
            }

            video.onended = () => {
                if (videoLoopRemaining === Infinity) {
                    // will not happen if video.loop=true, but keep safe
                    video.play().catch(e => console.error(e));
                    return;
                }

                videoLoopRemaining -= 1;
                if (videoLoopRemaining > 0) {
                    // replay video
                    video.currentTime = 0;
                    video.play().catch(err => console.error('replay error', err));
                } else {
                    notifyPython({ event: 'finished', type: 'video', src: item.video, index: currentIndex });
                    advanceToNext();
                }
            };

            video.src = item.video;
            video.muted = false;
            video.volume = 1;
            video.play().catch(err => console.error('video play error', err));
            notifyPython({ event: 'started', type: 'video', src: item.video, index: currentIndex });
        } else {
            // unknown item, skip
            console.warn('unknown playlist item', item);
            advanceToNext();
        }
    }

    function handlePlaylistMessage(newList) {
        // Replace old playlist immediately and start playing new list
        if (!Array.isArray(newList)) newList = [];
        playlist = newList.slice();
        if (playlist.length === 0) {
            // clear/stop
            stopAll();
            return;
        }

        // start from first item
        currentIndex = 0;
        startCurrentItem();
    }

    function pausePlayback() {
        if (isPaused) return; // already paused
        isPaused = true;

        // pause video
        try {
            if (!video.paused) video.pause();
        } catch (e) {}

        // pause image timer: compute remaining
        if (imageTimer) {
            const elapsed = Date.now() - imageStartAt;
            imageRemaining = Math.max(0, imageRemaining - elapsed);
            clearTimeout(imageTimer);
            imageTimer = null;
        }
    }

    function resumePlayback() {
        if (!isPaused) return;
        isPaused = false;

        // resume video
        try {
            if (video.src && video.paused) video.play().catch(e => console.error('resume play err', e));
        } catch (e) {}

        // resume image timer
        if (playingItem && playingItem.image && imageRemaining !== Infinity && imageRemaining > 0) {
            imageStartAt = Date.now();
            imageTimer = setTimeout(() => {
                notifyPython({ event: 'finished', type: 'image', src: playingItem.image, index: currentIndex });
                imageTimer = null;
                advanceToNext();
            }, imageRemaining);
        }
    }

    function parseControllerMessage(message) {
        try {
            console.log('Received from Monitor:', message);

            // If it's already an object (e.g. json parsed by the WS layer), return it
            if (message && typeof message === 'object' && !((typeof Buffer !== 'undefined') && Buffer.isBuffer(message)) && !(message instanceof ArrayBuffer) && !(message instanceof Uint8Array)) {
                return message;
            }

            // Node Buffer -> string
            if (typeof Buffer !== 'undefined' && Buffer.isBuffer(message)) {
                const s = message.toString('utf8');
                return JSON.parse(s);
            }

            // ArrayBuffer or TypedArray
            if (message instanceof ArrayBuffer) {
                const s = new TextDecoder().decode(new Uint8Array(message));
                return JSON.parse(s);
            }

            if (message instanceof Uint8Array) {
                const s = new TextDecoder().decode(message);
                return JSON.parse(s);
            }

            // Fallback: convert to string and parse
            const str = (typeof message === 'string') ? message : String(message);
            return JSON.parse(str);
        } catch (e) {
            console.error('parseControllerMessage error', e, message);
            return null;
        }
    }

    function handleControllerMessage(msg) {
        if (!msg || !msg.tasks) return;

        const t = msg.tasks;
        if (t === 'playlist') {
            // replace playlist immediately
            handlePlaylistMessage(msg.playlist || []);
            return;
        }

        if (t === 'pause') {
            pausePlayback();
            return;
        }

        if (t === 'play') {
            resumePlayback();
            return;
        }

        console.warn('unknown tasks', t);
    }

    // ---- 连接 WebSocket ----
    window.pptWS.connect("ws://localhost:8765", (message) => {
        const parsed = parseControllerMessage(message);
        handleControllerMessage(parsed);
    });

    // 启动时进入 idle
    setTimeout(() => {
        notifyPython({ event: "ready" });
    }, 500);
});
