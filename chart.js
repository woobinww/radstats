document.addEventListener('DOMContentLoaded', () => {
  const ctx = document.getElementById('chart').getContext('2d');
  let chart;
  
  // 시작일, 종료일 초기값 설정 (기본: 최근 1 개월)
  const endInput = document.getElementById('end-date');
  const startInput = document.getElementById('start-date');
  const today = new Date();
  const oneYearAgo = new Date();
  oneYearAgo.setMonth(oneYearAgo.getMonth() - 12);

  endInput.value = today.toISOString().slice(0, 10);
  startInput.value = oneYearAgo.toISOString().slice(0, 10);

  async function updateChart() {
    const xType = document.getElementById('xType').value;
    const device = document.getElementById('device').value;
    const startDate = startInput.value;
    const endDate = endInput.value;

    // 유효성 검사
    if (startDate && endDate && startDate > endDate) {
      alert('⚠️ 시작일이 종료일보다 늦을 수 없습니다.');
      return;
    }

    try {
      const result = await window.api.getChartData(xType, device, startDate, endDate);
      const labels = result.labels;
      const data = result.data;

      if (!labels.length || !data.length) {
        alert('데이터가 없습니다.');
        return;
      }

      if (chart) {
        chart.destroy();
      }

      chart = new Chart(ctx, {
        type: "line",
        data: {
          labels: labels,
          datasets: [
            {
              label: `${device} 검사 건수 (${xType})`,
              data: data,
              borderWidth: 2,
              borderColor: 'blue',
              fill: false,
              tension: 0.1,
            },
          ],
        },
        options: {
          responsive: true,
          plugins: {
            legend: {
              position: "top",
            },
            title: {
              display: true,
              text: "검사 건수 추이"
            }
          },
          scales: {
            x: {
              title: {
                display: true,
                text: xType === 'day' ? '일별' : xType === 'week' ? '주별' : '월별'
              }
            },
            y: {
              beginAtZero: true,
              title: {
                display: true,
                text: '검사 건수'
              },
              ticks: {
                autoSkip: true,
                maxTicksLimit: 50
              }
            },
          },
        },
      });
    } catch (e) {
      alert('차트 데이터를 불러오는 중 오류 발생');
      console.error('차트 데이터 불러오기 실패: ', e);
    }
  }

  document.getElementById('xType').addEventListener('change', updateChart);
  document.getElementById('device').addEventListener('change', updateChart);
  startInput.addEventListener('change', updateChart);
  endInput.addEventListener('change', updateChart);
  document.getElementById('showchart').addEventListener('click', updateChart);

  updateChart(); // 초기 실행
});
