const { app, BrowserWindow } = require("electron");
const path = require("path");

let mainWindow;

app.whenReady().then(() => {
    mainWindow = new BrowserWindow({
        width: 200,
        height: 360,
        x: 1200, // 屏幕右侧
        y: 600,
        frame: false,
        transparent: true,
        alwaysOnTop: true,
        skipTaskbar: true,
        webPreferences: {
            preload: path.join(__dirname, "preload.js"),
            contextIsolation: true,                      // 必须开启
            nodeIntegration: false                      // 必须关闭
        }
    });
    // 监听控制台消息
    mainWindow.webContents.on('console-message', (event, level, message, line, sourceId) => {
        console.log(`console: [${level}]: ${message} (${sourceId}:${line})`)
    });

    mainWindow.setAlwaysOnTop(true, "screen-saver");
    mainWindow.setIgnoreMouseEvents(false); // 可穿透可调整
    mainWindow.loadFile("index.html");
    // 临时打开 DevTools 以便调试 renderer 日志和 WebSocket 行为
    // try {
    //     mainWindow.webContents.openDevTools();
    // } catch (e) {
    //     console.warn('openDevTools failed:', e);
    // }
    console.log("Electron app is ready.");
});

app.on("window-all-closed", () => {
    if (process.platform !== "darwin") app.quit();
});
