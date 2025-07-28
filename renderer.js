// 버튼 클릭 시 처리

// 통계 생성
document.getElementById('generate-excel').addEventListener('click', async () => {
  const month = document.getElementById('target-month').value;
  const resultTag = document.getElementById('excel-result');

  if (!month) {
    resultTag.textContent = '⚠️ 생성할 월을 선택하세요.';
    return;
  }

  if (!/^\d{4}-\d{2}$/.test(month)) {
    resultTag.textContent = '⚠️ 형식: YYYY-MM으로 입력해주세요.';
    return;
  }

  resultTag.textContent = '⌛ 통계 생성 중...';

  try {
    const result = await window.api.writeStatistics(month); // <- month 전달
    resultTag.textContent = result;
  } catch (e) {
    resultTag.textContent = `❌ 오류: ${e.message}`;
  }
});

// xls -> csv -> sqlite db 저장 run all import
document.getElementById('run-all-import').addEventListener('click', async () => {
  const resultTag = document.getElementById('import-result');
  resultTag.textContent = '⌛ 파일 변환 및 DB 저장 중...';

  try {
    const convertMsg = await window.api.convertXlsToCsv();
    const saveMsg = await window.api.saveToDB();
    resultTag.textContent = `${convertMsg}\n${saveMsg}`;
  } catch (e) {
    resultTag.textContent = `❌ 오류 발생: ${e.message}`;
  }
});

// DOM에 버전 넣기
document.addEventListener('DOMContentLoaded', () => {
  const versionTag = document.getElementById('app-version');
  if (window.api?.getAppVersion && versionTag) {
    versionTag.textContent = `v${window.api.getAppVersion()}`;
  }
});


// 폴더 열기
document.getElementById('open-rawdata').addEventListener('click', () => {
  try {
    window.api.openFolder('data/raw_xls');
  } catch (e) {
    document.getElementById('folder-result').textContent = `❌ 폴더 열기 실패: ${e.message}`;
  }
});

document.getElementById('open-reports').addEventListener('click', () => {
  try {
    window.api.openFolder('reports');
  } catch (e) {
    document.getElementById('folder-result').textContent = `❌ 폴더 열기 실패: ${e.message}`;
  }
});
