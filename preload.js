const { contextBridge } = require('electron');
const path = require('path');
const fs = require('fs');

const { convertXlsToCsv } = require('./src/xls_to_csv.js');
const { saveToDB } = require('./src/save_to_sqlite.js');
const { writeStatistics } = require('./src/write_stats_to_excel');
const { getChartData } = require('./src/chart_data');

// app.getVersion() 은 main process에서만 동작. require 대신 package.json 직접 참조
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
      const result = await writeStatistics(targetMonth); // 인자 전달
      return result;
    } catch(e) {
      return `❌ 오류: ${e.message}`;
    }
  },
  convertXlsToCsv,
  saveToDB,
  getChartData: async (xType, device, startDate, endDate) => {
    try {
      return await getChartData(xType, device, startDate, endDate);
    } catch (e) {
      throw new Error(`차트 데이터 오류: ${e.message}`);
    }
  },
  getAppVersion: () => version
});
