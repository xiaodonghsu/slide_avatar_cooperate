const { app, BrowserWindow } = require("electron");
const path = require("path");

// 尝试加载 electron-log，若缺失则提供回退实现以避免主进程崩溃
let log;
try {
    log = require('electron-log');
} catch (e) {
    // minimal fallback that writes to console and exposes transports.file.getFile().path
    const { join } = require('path');
    const rootDir = path.join(__dirname, '..');
    const logDir = join(rootDir, 'log');
    try {
        const fs = require('fs');
        if (!fs.existsSync(logDir)) fs.mkdirSync(logDir);
    } catch (ee) {
        console.warn('create log dir failed', ee);
    }
    log = {
        info: (...args) => console.log(...args),
        warn: (...args) => console.warn(...args),
        error: (...args) => console.error(...args),
        debug: (...args) => console.debug(...args),
        transports: {
            file: {
                resolvePath: () => join(logDir, 'electron.log'),
                getFile: () => ({ path: join(logDir, 'electron.log') })
            }
        }
    };
}

// Configure log directory to project `log` folder when electron-log is available
try {
    const { join } = require('path');
    const rootDir = path.join(__dirname, '..');
    const logDir = join(rootDir, 'log');
    const fs = require('fs');
    if (!fs.existsSync(logDir)) fs.mkdirSync(logDir);
    if (log && log.transports && log.transports.file && typeof log.transports.file.resolvePath === 'function') {
        log.transports.file.resolvePath = () => join(logDir, 'electron.log');
    }
    log.info(logDir);
    const logPath = (log && log.transports && log.transports.file && typeof log.transports.file.getFile === 'function') ? (log.transports.file.getFile().path) : logDir;
    try { log.info('Logger initialized, writing to', logPath); } catch (_) { console.log('Logger initialized, writing to', logPath); }
} catch (e) {
    console.warn('logger init failed', e);
}

let mainWindow;

app.whenReady().then(() => {
    mainWindow = new BrowserWindow({
        width: 200,
        height: 360,
        x: 1100, // 屏幕右侧
        y: 380,
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
    // 监听控制台消息并写入 electron-log
    mainWindow.webContents.on('console-message', (event, level, message, line, sourceId) => {
        log.info(`console: [${level}]: ${message} (${sourceId}:${line})`);
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
    log.info("Electron app is ready.");
});

app.on("window-all-closed", () => {
    if (process.platform !== "darwin") app.quit();
});
