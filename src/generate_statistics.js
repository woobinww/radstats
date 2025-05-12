const sqlite3 = require("sqlite3").verbose();
const path = require("path");

// DB 연결
const dbPath = path.join(__dirname, "../db/database.sqlite");
const db = new sqlite3.Database(dbPath);

// 결과 저장용 객체
const stats = {};

// 쿼리: 월별 장비/부서별 건수 추출
const query = `
  SELECT
    substr(검사시간, 1, 7) AS month, -- 'YYYY-MM' 형식으로 자름
    장비,
    부서,
    COUNT(*) AS count
  FROM rad_exam
  WHERE NOT (
    (장비 = 'MR' AND 검사명 LIKE '%외부%') -- 외부 MRI 제외 조건
    OR
    (장비 = 'US' AND 검사명 LIKE '%상담%') -- 외부 MRI 제외 조건
  )
  GROUP BY month, 장비, 부서
  ORDER BY month ASC
`;

db.all(query, [], (err, rows) => {
  if (err) {
    console.error("❌ 쿼리 오류:", err.message);
    return;
  }

  for (const row of rows) {
    const { month, 장비, 부서, count } = row;

    // 월별 구조가 없으면 생성 
    if (!stats[month]) {
      stats[month] = { 장비: {}, 부서: {} };
    }

    // 장비별 누적
    stats[month]["장비"][장비] = (stats[month]["장비"][장비] || 0) + count;

    // 부서별 누적
    stats[month]["부서"][부서] = (stats[month]["부서"][부서] || 0) + count;
  }

  console.log("✅ 통계 집계 완료:");
  console.log(JSON.stringify(stats, null, 2)); // 보기 좋게 출력

  db.close();
});
