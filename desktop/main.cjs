const { app, BrowserWindow, screen, globalShortcut, ipcMain } = require("electron");
const path = require("node:path");

let mainWindow;
let timerWindow;

function createWindows() {
  mainWindow = new BrowserWindow({ width: 1440, height: 920, webPreferences: { contextIsolation: true } });
  mainWindow.loadFile(path.join(__dirname, "..", "index.html"));
  ipcMain.handle("set-always-on-top", (_event, enabled) => {
    if (!timerWindow) return false;
    timerWindow.setAlwaysOnTop(Boolean(enabled), "floating");
    return timerWindow.isAlwaysOnTop();
  });
  ipcMain.handle("display-bounds", () => screen.getAllDisplays().map((display) => ({ id: display.id, bounds: display.bounds, label: display.label })));
}

app.whenReady().then(() => {
  createWindows();
  globalShortcut.register("CommandOrControl+Shift+T", () => {
    if (!timerWindow) {
      timerWindow = new BrowserWindow({ width: 260, height: 105, minWidth: 150, minHeight: 70, transparent: true, frame: false, alwaysOnTop: true, resizable: true, skipTaskbar: true });
      timerWindow.loadFile(path.join(__dirname, "timer.html"));
    }
    timerWindow.show();
  });
});
app.on("will-quit", () => globalShortcut.unregisterAll());
