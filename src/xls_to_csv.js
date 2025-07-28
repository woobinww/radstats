const xlsx = require("xlsx");
const fs = require("fs");
const path = require("path");

/*
  * xls 또는 xlsx 파일을 csv로 변환
  * @param {string} userDataPath - 사용자 데이터 루트 경로
  * @return {string} - 변환 결과 메세지
*/

function convertXlsToCsv(userDataPath) {
  const rawFolder = path.join(userDataPath, "data", "raw_xls");
  const convertedFolder = path.join(userDataPath, "data", "converted_csv");
  const results = [];

  // 폴더 없으면 생성
  fs.mkdirSync(rawFolder, { recursive: true })
  fs.mkdirSync(convertedFolder, { recursive: true })

  // .xls 파일을 순회하며 .csv로 변환
  fs.readdirSync(rawFolder).forEach((file) => {
    if (file.endsWith(".xls") || file.endsWith(".xlsx")) {
      const outputName = file.replace(/\.(xls|xlsx)$/, ".csv");
      const outputPath = path.join(convertedFolder, outputName);
  
      // 이미 변환된 파일은 건너뜀
      if (fs.existsSync(outputPath)) {
        console.log(`⏩ 이미 변환됨: ${outputName} → 건너뜁니다.`);
        return;
      }
  
      const filePath = path.join(rawFolder, file);
      const workbook = xlsx.readFile(filePath);
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const csv = xlsx.utils.sheet_to_csv(sheet);
  
      fs.writeFileSync(outputPath, csv);
      results.push(`✅ 변환 완료: ${file} → ${outputName}`);
    }
  });

  return results.join('\n');
}

module.exports = { convertXlsToCsv };
