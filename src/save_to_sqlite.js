const sqlite3 = require("sqlite3").verbose();
const fs = require("fs");
const path = require("path");
const { parse } = require("csv-parse/sync");

function saveToDB(userDataPath) {
  return new Promise((resolve, reject) => {
    
    // db 연결
    const dbPath = path.join(userDataPath, "db", "database.sqlite");
    fs.mkdirSync(path.dirname(dbPath), { recursive: true })

    const db = new sqlite3.Database(dbPath, (err) => {
      if (err) return reject(err);
    });
    
    // 테이블 생성 (존재하면 무시)
    db.serialize(() => {
      db.run(`
        CREATE TABLE IF NOT EXISTS rad_exam (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          환자번호 TEXT,
          환자명 TEXT,
          나이성별 TEXT,
          장비 TEXT,
          검사부위 TEXT,
          검사명 TEXT,
          부서 TEXT,
          처방의 TEXT,
          병동 TEXT,
          검사시간 TEXT,
          병원 TEXT,
          AccessNo TEXT
        )
      `, (err) => { if (err) return reject(err); });
    
      // 가져왔던 파일 import_log 에 추가
      db.run(`
        CREATE TABLE IF NOT EXISTS import_log (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          filename TEXT UNIQUE,
          imported_at TEXT
        )
      `, (err) => { if (err) return  reject(err); });
    });
    
    // CSV 파일 읽기
    const csvFolder = path.join(userDataPath, "data", "converted_csv");
    fs.mkdirSync(csvFolder, { recursive: true })
    // 가져온 파일 목록 확인
    const files = fs.readdirSync(csvFolder).filter(f => f.endsWith(".csv"));
    // 파일 카운트
    let savedCount = 0;
    let skippedCount = 0;
    
    // CSV 파일 처리
    function processCsvFileAtIndex(index) {
      if (index >= files.length) {
        db.close((err) => {
          if (err) return reject(err);
          resolve(`✅ DB 저장 완료\n신규: ${savedCount}건, 중복: ${skippedCount}건`);
        });
        return;
      }
    
      const file = files[index];
    
      // 처리 여부 확인
      db.get("SELECT * FROM import_log WHERE filename = ?", [file], (err, row) => {
        if (err) return reject(err); // 오류가 발생되면 reject. 디버깅 용이

        if (row) {
          skippedCount++;
          processCsvFileAtIndex(index + 1);
        } else {
          try {
            // CSV 읽기 및 DB 저장
            const fullPath = path.join(csvFolder, file);
            const content = fs.readFileSync(fullPath);
            const records = parse(content, { columns: true, skip_empty_lines: true });
          
            const stmt = db.prepare(`
              INSERT INTO rad_exam (
                환자번호, 환자명, 나이성별, 장비, 검사부위, 검사명,
                부서, 처방의, 병동, 검사시간, 병원, AccessNo
              ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            `, (err) => { if (err) return reject(err); });
        
            for (const rec of records) {
              stmt.run(
                rec['환자번호'],
                rec['환자명'],
                rec['나이/성별'],
                rec['장비'],
                rec['검사부위'],
                rec['검사명'],
                rec['부서'],
                rec['처방의'],
                rec['병동'],
                rec['Time'],
                rec['병원'],
                rec['Access No.'],
                (err) => { if (err) return reject(err); }
              );
            }
            stmt.finalize((err) => {
              if (err) return reject(err);
  
              // 처리한 파일 기록
              db.run(
                "INSERT INTO import_log (filename, imported_at) VALUES (?, datetime('now'))", 
                [file],
                (err) => {
                  if (err) return reject(err);
                  savedCount++;
                  processCsvFileAtIndex(index + 1);
                }
              );
            });
          } catch (e) {
            return reject(e);
          }
        }
      });
    }
    
    processCsvFileAtIndex(0);
  });
}

module.exports = { saveToDB };
