const { app, BrowserWindow } = require('electron');
const fs = require('fs');
const path = require('path');

// 사용자 데이터 경로
const userData = app.getPath('userData');

// 개발 환경에서만 NODE_ENV 설정
if (!app.isPackaged) {
  process.env.NODE_ENV = 'development';
}

function ensureFolders() {
  const folders = [
    'db',
    path.join('data', 'raw_xls'),
    path.join('data', 'converted_csv'),
    'reports'
  ];

  folders.forEach((relPath) => {
    const fullPath = path.join(userData, relPath);
    if (!fs.existsSync(fullPath)) {
      fs.mkdirSync(fullPath, { recursive: true });
      console.log(`📁 폴더 생성됨: ${fullPath}`);
    }
  });
}

function createWindow() {
  const win = new BrowserWindow({
    width: 800,
    height: 1200,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,  // 반드시 true (보안상)
      enableRemoteModule: false,
      nodeIntegration: false,
      sandbox: false,           // electron v12+ 에서는 false 권장
      additionalArguments: [
        `--userData=${app.getPath('userData')}`,
        `--env=${process.env.NODE_ENV || 'production'}`
      ] // preload.js에 경로전달
    }
  });

  win.loadFile('index.html');

  win.webContents.on('did-fail-load', (event, errorCode, errorDesc) => {
    console.error(`❌ 페이지 로딩 실패: ${errorCode} - ${errorDesc}`)
  });
  // ✅ preload 실행 중 오류 발생 확인
  win.webContents.on('console-message', (e, level, message) => {
    console.log(`🖥️ [preload console]: ${message}`);
  });
}

app.whenReady().then(() => {
  ensureFolders();
  createWindow();
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});
