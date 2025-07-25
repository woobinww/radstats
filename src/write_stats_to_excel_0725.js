
const ExcelJS = require('exceljs');
const sqlite3 = require('sqlite3').verbose();
const path = require('path');
const fs = require('fs');

// 입력할 대상 월
// const targetMonth = '2025-04';

// DB 및 템플릿 파일 경로
const dbPath = path.join(__dirname, '../db/database.sqlite');
const templatePath = path.join(__dirname, '../templates/stat_form.xlsx');
// 저장할 파일 경로 없으면 자동 생성
const reportsDir = path.join(__dirname, '../reports');
if (!fs.existsSync(reportsDir)) fs.mkdirSync(reportsDir);


// 조건별 셀 매핑 (검사명 키워드는 없으면 null, 있으면 OR조건으로 사용)
const doctors = ['기세린', '강진우', '유재중', '김태겸', '김건우'];
const wards = ['외래', '기타'];
const modalities = ['CR', 'US', 'CT', 'MR', 'RF'];
const colLetters = 'BCDEFGHIJKLMN'.split('');

const sheetMaps = {};

// 총 통계 sheet Mapping
sheetMaps['총 통계'] = {};

const getCell = (doctorIndex, wardIndex, row) =>
  `${colLetters[doctorIndex * 2 + wardIndex]}${row}`;

const getWildcardCell = (row) =>
  `${colLetters[doctors.length * 2]}${row}`;

// CR, CT, MR, RF
const baseRowByModality = {
  CR: 7,
  CT: 11,
  MR: 12,
  RF: 13,
};

modalities.forEach(modality => {
  if (modality === 'US') return; // US는 따로 처리
  const row = baseRowByModality[modality];
  doctors.forEach((doctor, i) => {
    wards.forEach((ward, j) => {
      const key = `${modality}|${doctor}|${ward}`;
      sheetMaps['총통계'][key] = getCell(i, j, row);
    });
  });
  sheetMaps['총 통계'][`${modality}|*|*`] = getWildcardCell(row);
});

// CR age
const ageRow = 14;
doctors.forEach((doctor, i) => {
  wards.forEach((ward, j) => {
    const key = `CR|${doctor}|${ward}|age`;
    sheetMaps['총 통계'][key] = getCell(i, j, ageRow);
  });
});
sheetMaps['총 통계']['CR|*|*|age'] = getWildcardCell(ageRow);

// US 전용 분리 처리
const usRows = {
  abdo: 8,
  doppler: 9,
  etc: 10
};

// US - 복부
doctors.forEach((doctor, i) => {
  wards.forEach((ward, j) => {
    const key = `US|${doctor}|${ward}|복부|abdo`;
    sheetMaps['총 통계'][key] = getCell(i, j, usRows.abdo);
  });
});
sheetMaps['총 통계']['US|*|*|복부|abdo'] = getWildcardCell(usRows.abdo);

// US - Doppler
doctors.forEach((doctor, i) => {
  wards.forEach((ward, j) => {
    const key = `US|${doctor}|${ward}|Doppler|Dopper`;
    sheetMaps['총 통계'][key] = getCell(i, j, usRows.doppler);
  });
});
sheetMaps['총 통계']['US|*|*|Doppler|Dopper'] = getWildcardCell(usRows.doppler);

// US - 기타 (!복부 && !doppler)
doctors.forEach((doctor, i) => {
  wards.forEach((ward, j) => {
    const key = `US|${doctor}|${ward}|!복부|!abdo|!Doppler|!dopper`;
    sheetMaps['총 통계'][key] = getCell(i, j, usRows.etc);
  });
});
sheetMaps['총 통계']['US|*|*|!복부|!abdo|!Doppler|!dopper'] = getWildcardCell(usRows.etc);

// 결과 확인
console.log(sheetMaps['총 통계']);









// "항목별 통계" 시트 타겟 월 구성
function generateLastYearToCurrentMonths(targetMonth) {
  const [targetYearStr, targetMonStr] = targetMonth.split('-').map(Number);
  const currentYear = parseInt(targetYearStr, 10);
  const lastYear = currentYear - 1;
  const months = [];

  for (let m = 1; m <= 12; m++) {
    months.push(`${lastYear}-${String(m).padStart(2, '0')}`);
  }
  for (let m = 1; m <= parseInt(targetMonStr, 10); m++) {
    months.push(`${currentYear}-${String(m).padStart(2, '0')}`);
  }
  return { months, lastYear, currentYear };
}

// "항목별 통계" 시트의 해당 자리 알파벳 인덱스
function getExcelColumnLetter(index) {
  const base = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  let result = '';
  while (index >= 0) {
    result = base[index % 26] + result;
    index = Math.floor(index / 26) - 1;
  }
  return result;
}

//  항목별 통계 시트용 rowMap 생성 함수
function createStatRowMap(lastYear, currentYear) {
  return {
    'CR': { rowMap: { [lastYear]: 4, [currentYear]: 5 }, startCol: 'B' },
    'CT': { rowMap: { [lastYear]: 10, [currentYear]: 11 }, startCol: 'B' },
    'MR': { rowMap: { [lastYear]: 16, [currentYear]: 17 }, startCol: 'B' },
    'US': { rowMap: { [lastYear]: 22, [currentYear]: 23 }, startCol: 'B' },
  };
}



// 메인 함수
async function writeStatistics(targetMonth) {
  const db = new sqlite3.Database(dbPath);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(templatePath);

  const [year, month] = targetMonth.split('-');
  const outputPath = path.join(
    __dirname,
    `../reports/${year}_${month}_영상의학과_월간통계.xlsx`
  );

  for (const sheetName in sheetMaps) {
    const sheet = workbook.getWorksheet(sheetName);
    const cellMap = sheetMaps[sheetName];

    if (!sheet) continue; // sheet 존재하지 않는 경우 넘어감. 오류 방지

    // "총 통계" 시트일 때만 A4에 날짜 표시
    if (sheetName === "총 통계") {
      sheet.getCell('A4').value = `${year}년 ${month}월`;
    }

    if (sheetName === "MRI") {
      sheet.getCell('A1').value = `의료진별 MRI 통계(${year}년 ${month}월)`;
    }
    if (sheetName === "CT") {
      sheet.getCell('A1').value = `진료과별 CT 통계(${year}년 ${month}월)`;
    }

    for (const key in cellMap) {
      const [장비, 처방의, 병동, ...검사명키워드] = key.split('|');
      const cell = cellMap[key];
  
      const whereClauses = [
        "REPLACE(substr(검사시간, 1, 7), '/', '-') = ?",
        "TRIM(장비) = ?", // TRIM : 문자열 앞뒤의 공백 제거
      ];
      const params = [targetMonth, 장비];
  
      if (처방의 && 처방의 !== '*') {
        whereClauses.push("TRIM(처방의) = ?");
        params.push(처방의);
      }
  
      if (병동 && 병동 !== '*') {
        if (병동 === '기타' || 병동.startsWith('!')) {
          const exclude = 병동 === '기타' ? '외래' : 병동.slice(1);
          whereClauses.push("TRIM(병동) != ?");
          params.push(exclude);
        } else {
          whereClauses.push("TRIM(병동) = ?");
          params.push(병동);
        }
      }
  
      // 기본 제외 조건
      whereClauses.push("NOT (장비 = 'MR' AND 검사명 LIKE '%외부%')");
      whereClauses.push("NOT (장비 = 'US' AND 검사명 LIKE '%상담%')");
      whereClauses.push("NOT (장비 = 'RF' AND 검사명 NOT LIKE '%외래%')");
  
      // 검사명 필터 처리
      if (검사명키워드.length > 0) {
        const includeRaw = 검사명키워드.filter(k => !k.startsWith('!'));
        const exclude = 검사명키워드.filter(k => k.startsWith('!')).map(k => k.slice(1));

        const orKeywords = [];
        const andKeywords = [];

        for (const word of includeRaw) {
          if (word.includes('+')) {
            andKeywords.push(...word.split('+'));
          } else {
            orKeywords.push(word);
          }
        }
  
        if (orKeywords.length > 0) {
          whereClauses.push(
            `(${orKeywords.map(() => '검사명 LIKE ? COLLATE NOCASE').join(' OR ')})`
          );
          params.push(...orKeywords.map(k => `%${k}%`));
        }

        if (andKeywords.length > 0) {
          whereClauses.push(
            `${andKeywords.map(() => '검사명 LIKE ? COLLATE NOCASE').join(' AND ')}`
          );
          params.push(...andKeywords.map(k => `%${k}%`));
        }

        if (exclude.length > 0) {
          whereClauses.push(
            `${exclude.map(() => '검사명 NOT LIKE ? COLLATE NOCASE').join(' AND ')}`
          );
          params.push(...exclude.map(k => `%${k}%`));
        }
      }
  
      const query = `
        SELECT COUNT(*) AS count
        FROM rad_exam
        WHERE ${whereClauses.join(' AND ')}
      `;
  
      const count = await new Promise((resolve, reject) => {
        db.get(query, params, (err, row) => {
          if (err) reject(err);
          else resolve(row.count);
        });
      });
  
      sheet.getCell(cell).value = count;
    }
  }

  // '항목별 통계' 시트 처리
  const { months: monthList, lastYear, currentYear } = generateLastYearToCurrentMonths(targetMonth);
  const 항목별Sheet = workbook.getWorksheet('항목별 통계');
  const 항목별통계Map = createStatRowMap(lastYear, currentYear)
  if (항목별Sheet) {
    for (const 장비 in 항목별통계Map) {
      for (const month of monthList) {
        const [y, m] = month.split('-');
        const col = getExcelColumnLetter(m - 1 + 1); // B 부터 시작
        const row = 항목별통계Map[장비].rowMap[y];
        if (!row) continue;
        const cell = `${col}${row}`;

        const count = await new Promise((resolve, reject) => {
          const where = [
            "substr(REPLACE(검사시간, '/', '-'), 1, 7) = ?",
            "TRIM(장비) = ?"
          ];
          const params = [month, 장비];

          if (장비 == 'MR') {
            where.push("검사명 NOT LIKE '%외부%'");
          }

          if (장비 == 'US') {
            where.push("검사명 NOT LIKE '%상담%'");
          }
          
          const query = `
            SELECT COUNT(*) AS count
            FROM rad_exam
            WHERE ${where.join(' AND ')}
          `;

          db.get(query, params, (err, row) => {
            if (err) reject(err);
            else resolve(row.count);
          });
        });

        항목별Sheet.getCell(cell).value = count;
      }
    }
  }

  await workbook.xlsx.writeFile(outputPath);
  db.close();
  console.log(`✅ ${targetMonth} 엑셀 저장 완료 → ${outputPath}`);
}

module.exports = { writeStatistics };


