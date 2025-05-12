const xlsx = require("xlsx");
const fs = require("fs");
const path = require("path");

// 변환할 폴더 경로
const rawFolder = path.join(__dirname, "../data/raw_xls");
const convertedFolder = path.join(__dirname, "../data/converted_csv");

function convertXlsToCsv() {
  const results = [];

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
