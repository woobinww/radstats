const { app, BrowserWindow } = require('electron');
const fs = require('fs');
const path = require('path');

// 필요한 폴더 생성
const requireDirs = ['db', 'data/raw_xls', 'data/converted_csv', 'reports'];
requireDirs.forEach(dir => {
  const fullPath = path.join(__dirname, dir);
  if (!fs.existsSync(fullPath)) {
    fs.mkdirSync(fullPath, { recursive: true });
    console.log(`📁 폴더 생성됨: %{dir}`);
  }
});

function createWindow() {
  const win = new BrowserWindow({
    width: 800,
    height: 1200,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,  // 반드시 true (보안상)
      enableRemoteModule: false,
      nodeIntegration: false,
      sandbox: false           // electron v12+ 에서는 false 권장
    }
  });

  win.loadFile('index.html');
}

app.whenReady().then(createWindow);

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});
