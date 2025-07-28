const { contextBridge, shell } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');

// 개발 환경에서만 preload 로깅 개발 환경 여부 확인
const isDev = process.argv.some(arg => arg === '--env=development');

if (isDev) {
  console.log('✅ preload.js 시작됨 (개발모드)');
}

// userData 경로 설정
let userData;

// 우선순위: 1. 추가인자 -> 2. 개발환경 경로 -> 3. 배포 기본 경로
if (isDev) {
  userData = path.join(__dirname);
} else {
  // 배포 환경에서는 --userData 인자가 오면 사용, 없으면 기본 AppData 경로사용
  const userDataArg = process.argv.find(arg => arg.startsWith('--userData='));
  if (userDataArg) {
    const extracted = userDataArg.split('=')[1];
    if (fs.existsSync(extracted)) {
      userData = extracted;
    }
  }

  if (!userData) {
    userData = path.join(os.homedir(), 'AppData', 'Roaming', 'RadStats');
  }
}

console.log('NODE_ENV:', process.env.NODE_ENV);
console.log('isDev:', isDev);
console.log('userData:', userData);

const { convertXlsToCsv } = require('./src/xls_to_csv.js');
const { saveToDB } = require('./src/save_to_sqlite.js');
const { writeStatistics } = require('./src/write_stats_to_excel');
const { getChartData } = require('./src/chart_data');

// app.getVersion() 은 main process에서만 동작. require 대신 package.json 직접 참조
// 버전정도 가져오기 
const packageJsonPath = path.join(__dirname, 'package.json');
let version = 'unknown';

try {
  const packageData = JSON.parse(fs.readFileSync(packageJsonPath, 'utf-8'));
  version = packageData.version || version;
} catch (e) {
  console.warn('⚠️ 버전 정보를 불러올 수 없습니다: ', e.message);
}

contextBridge.exposeInMainWorld('api', {
  writeStatistics: async (targetMonth) => {
    try {
      const result = await writeStatistics(targetMonth, userData); // 인자 전달
      return result;
    } catch(e) {
      return `❌ 오류: ${e.message}`;
    }
  },
  convertXlsToCsv: () => {
    return convertXlsToCsv(userData); // userData는 preload에서 정의된 경로
  },
  saveToDB: async () => {
    try {
      return await saveToDB(userData);
    } catch (e) {
      return `❌ DB 저장 실패: ${e.message}`;
    }
  },
  getChartData: async (xType, device, startDate, endDate) => {
    try {
      return await getChartData(xType, device, startDate, endDate, userData);
    } catch (e) {
      throw new Error(`차트 데이터 오류: ${e.message}`);
    }
  },
  getAppVersion: () => version,
  getUserPath: (...rel) => {
    const fullPath = path.join(userData, ...rel);
    fs.mkdirSync(path.dirname(fullPath), { recursive: true });
    return fullPath;
  },
  openFolder: (relPath) => {
    const fullPath = path.isAbsolute(relPath)
      ? relPath
      : path.join(userData, relPath);

    if (fs.existsSync(fullPath)) {
      shell.openPath(fullPath);
    } else {
      throw new Error('해당 경로가 존재하지 않습니다.');
    }
  }
});
