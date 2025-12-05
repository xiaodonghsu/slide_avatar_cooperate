const { contextBridge } = require("electron");

// 尝试加载 electron-log，若缺失则回退到 console 实现，防止 preload 加载失败
let _elog = null;
try {
    _elog = require('electron-log');
} catch (e) {
    // fallback: minimal shim that writes to console
    _elog = {
        info: (...args) => console.log(...args),
        warn: (...args) => console.warn(...args),
        error: (...args) => console.error(...args),
        debug: (...args) => console.debug(...args)
    };
}

// 将 electron-log 的简单接口暴露给渲染进程
contextBridge.exposeInMainWorld('electronLog', {
    info: (...args) => _elog.info(...args),
    warn: (...args) => _elog.warn(...args),
    error: (...args) => _elog.error(...args),
    debug: (...args) => _elog.debug(...args)
});

// 在 preload 中管理原生 WebSocket，避免将原生对象跨 contextIsolation 返回
contextBridge.exposeInMainWorld("pptWS", {
    connect(url, onMessage) {
        console.log("WebSocket -> Connecting...", url);

        // 默认重连配置
        const defaultOpts = {
            reconnect: true,
            initialDelay: 1000, // 1s
            maxDelay: 30000, // 30s
            factor: 1.5 // 指数退避倍数
        };

        // 清理已有连接并停止已存在的重连计时器
        try {
            if (globalThis._ppt_ws_reconnect_timer) {
                clearTimeout(globalThis._ppt_ws_reconnect_timer);
                globalThis._ppt_ws_reconnect_timer = null;
            }
            if (globalThis._ppt_ws) {
                try { globalThis._ppt_ws.close(); } catch (e) {}
                globalThis._ppt_ws = null;
            }
        } catch (e) {}

        // 初始化重连状态
        globalThis._ppt_ws_should_reconnect = true;
        globalThis._ppt_ws_reconnect_delay = defaultOpts.initialDelay;

        const createWebSocket = () => {
            try {
                const ws = new WebSocket(url);
                globalThis._ppt_ws = ws;

                ws.onopen = () => {
                    console.log("WebSocket -> Connected");
                    // 重置重连延迟
                    globalThis._ppt_ws_reconnect_delay = defaultOpts.initialDelay;
                    if (globalThis._ppt_ws_reconnect_timer) {
                        clearTimeout(globalThis._ppt_ws_reconnect_timer);
                        globalThis._ppt_ws_reconnect_timer = null;
                    }
                };

                ws.onerror = (err) => console.log("WebSocket Error:", err);

                ws.onclose = (ev) => {
                    console.log("WebSocket -> Disconnected", ev && ev.code);
                    // 如果被标记为需要重连，则持续尝试
                    if (globalThis._ppt_ws_should_reconnect && defaultOpts.reconnect) {
                        const delay = Math.min(globalThis._ppt_ws_reconnect_delay, defaultOpts.maxDelay);
                        console.log(`WebSocket -> Reconnecting in ${delay}ms`);
                        globalThis._ppt_ws_reconnect_timer = setTimeout(() => {
                            // 增加下次延迟
                            globalThis._ppt_ws_reconnect_delay = Math.min(Math.floor(globalThis._ppt_ws_reconnect_delay * defaultOpts.factor), defaultOpts.maxDelay);
                            createWebSocket();
                        }, delay);
                    }
                };

                ws.onmessage = (event) => {
                    try {
                        const msg = JSON.parse(event.data);
                        onMessage(msg);
                    } catch (e) {
                        console.warn("Accept None JSON message: ", event.data);
                    }
                };
            } catch (e) {
                console.error('createWebSocket error', e);
                // 若创建本身抛异常，尝试延迟重连
                if (globalThis._ppt_ws_should_reconnect) {
                    const delay = Math.min(globalThis._ppt_ws_reconnect_delay, defaultOpts.maxDelay);
                    globalThis._ppt_ws_reconnect_timer = setTimeout(() => {
                        globalThis._ppt_ws_reconnect_delay = Math.min(Math.floor(globalThis._ppt_ws_reconnect_delay * defaultOpts.factor), defaultOpts.maxDelay);
                        createWebSocket();
                    }, delay);
                }
            }
        };

        // 开始首次连接
        createWebSocket();
    },

    send(message) {
        try {
            const ws = globalThis._ppt_ws;
            if (ws && ws.readyState === WebSocket.OPEN) {
                ws.send(typeof message === 'string' ? message : JSON.stringify(message));
            } else {
                console.warn('pptWS.send: websocket not open');
            }
        } catch (e) {
            console.error('pptWS.send error', e);
        }
    },

    close() {
        try {
            // 停止自动重连并清理计时器
            globalThis._ppt_ws_should_reconnect = false;
            if (globalThis._ppt_ws_reconnect_timer) {
                clearTimeout(globalThis._ppt_ws_reconnect_timer);
                globalThis._ppt_ws_reconnect_timer = null;
            }
            if (globalThis._ppt_ws) {
                globalThis._ppt_ws.close();
                globalThis._ppt_ws = null;
            }
        } catch (e) {
            console.error('pptWS.close error', e);
        }
    }
});
