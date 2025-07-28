const sqlite3 = require('sqlite3').verbose();
const path = require('path');


function getChartData(xType, device, startDate, endDate, userDataPath) {
  return new Promise((resolve, reject) => {
    const dbPath = path.join(userDataPath, 'db', 'database.sqlite');
    const db = new sqlite3.Database(dbPath);

    let groupByClause = '';
    let dateColumn = `REPLACE(검사시간, '/', '-')`;

    switch (xType) {
      case 'day':
        groupByClause = `substr(${dateColumn}, 1, 10)`; // YYYY-MM-DD
        break;
      case 'week':
        groupByClause = `strftime('%Y-W%W', ${dateColumn})`; // YYYY-Wxx
        break;
      case 'month':
      default:
        groupByClause = `substr(${dateColumn}, 1, 7)`; // YYYY-MM
        break;
    }

    const query = `
      SELECT ${groupByClause} as label, COUNT(*) as count
      FROM rad_exam
      WHERE TRIM(장비) = ?
        AND ${dateColumn} BETWEEN ? AND ?
        AND NOT (장비 = 'MR' AND 검사명 LIKE '%외부%')
        AND NOT (장비 = 'US' AND 검사명 LIKE '%상담%')
      GROUP BY label
      ORDER BY label ASC
    `;

    db.all(query, [device, startDate, endDate], (err, rows) => {
      db.close();
      if (err) {
        reject(err);
      } else {
        const labels = rows.map(r => r.label);
        const data = rows.map(r => r.count);
        resolve({ labels, data });
      }
    });
  });
}

module.exports = { getChartData };
