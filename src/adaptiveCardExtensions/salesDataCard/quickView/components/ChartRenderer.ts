import { Chart, registerables } from 'chart.js';
import { ISalesDataSeries } from '../../../../common/models/models';

Chart.register(...registerables);

const MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

function formatDateLabel(d: Date): string {
  return `${MONTHS[d.getMonth()]} ${d.getDate()}`;
}

export class ChartRenderer {
  public static render(series: ISalesDataSeries[], daysBack: number): Promise<string> {
    return new Promise<string>((resolve) => {
      if (!series || series.length === 0) {
        resolve('');
        return;
      }

      const container = document.createElement('div');
      container.style.cssText = 'position:absolute;left:-9999px;top:-9999px;width:760px;height:380px';

      const canvas = document.createElement('canvas');
      canvas.width = 760;
      canvas.height = 380;
      container.appendChild(canvas);
      document.body.appendChild(container);

      try {
        const ctx = canvas.getContext('2d');
        if (!ctx) {
          resolve('');
          return;
        }

        const labels = (series[0]?.dataPoints ?? []).map(dp => formatDateLabel(dp.date));
        const pointRadius = (series[0]?.dataPoints.length ?? 0) <= 30 ? 3 : 1;

        const datasets = series.map(s => ({
          label: s.name,
          data: s.dataPoints.map(dp => dp.amount),
          borderColor: s.color,
          backgroundColor: s.color + '22',
          borderWidth: 2,
          pointRadius,
          tension: 0.3,
          fill: false
        }));

        const chart = new Chart(ctx, {
          type: 'line',
          data: { labels, datasets },
          options: {
            animation: false,
            responsive: false,
            plugins: {
              legend: {
                display: true,
                position: 'top',
                labels: { boxWidth: 10, font: { size: 10 }, padding: 8 }
              },
              tooltip: { enabled: false }
            },
            scales: {
              x: {
                ticks: { font: { size: 9 }, maxTicksLimit: 8, maxRotation: 0 },
                grid: { display: false }
              },
              y: {
                beginAtZero: true,
                ticks: {
                  font: { size: 9 },
                  callback: (val: number | string) => `$${Number(val).toLocaleString()}`
                }
              }
            }
          }
        });

        const dataUrl = canvas.toDataURL('image/png');
        chart.destroy();
        resolve(dataUrl);
      } finally {
        if (container.parentNode) {
          container.parentNode.removeChild(container);
        }
      }
    });
  }
}
