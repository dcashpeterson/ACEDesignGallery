import * as React from 'react';
import { Chart, registerables } from 'chart.js';
import { ISalesDataSeries } from '../../../common/models/models';
import * as strings from 'SalesBarChartCardAdaptiveCardExtensionStrings';

Chart.register(...registerables);

interface ISalesBarChartComponentProps {
  series: ISalesDataSeries[];
  totalSales: number;
  isLoading: boolean;
}

export const SalesBarChartComponent: React.FC<ISalesBarChartComponentProps> = ({ series, totalSales, isLoading }) => {
  const canvasRef = React.useRef<HTMLCanvasElement>(null);
  const chartRef = React.useRef<Chart | null>(null);

  React.useEffect(() => {
    if (!canvasRef.current || series.length === 0) return;

    const labels = (series[0]?.dataPoints ?? []).map(dp => {
      const d = dp.date instanceof Date ? dp.date : new Date(dp.date);
      return `${strings.Months[d.getMonth()]} ${d.getFullYear()}`;
    });

    const datasets = series.map(s => ({
      label: s.name,
      data: s.dataPoints.map(dp => dp.amount),
      backgroundColor: s.color + 'cc',
      borderColor: s.color,
      borderWidth: 1
    }));

    chartRef.current = new Chart(canvasRef.current, {
      type: 'bar',
      data: { labels, datasets },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            display: true,
            position: 'top',
            labels: { boxWidth: 12, font: { size: 11 }, padding: 10 }
          },
          tooltip: {
            callbacks: {
              label: ctx => `${ctx.dataset.label}: $${Number(ctx.raw).toLocaleString()}`
            }
          }
        },
        scales: {
          x: {
            ticks: { font: { size: 10 }, maxRotation: 45 },
            grid: { display: false }
          },
          y: {
            beginAtZero: true,
            ticks: {
              font: { size: 10 },
              callback: (val: number | string) => `$${Number(val).toLocaleString()}`
            }
          }
        }
      }
    });

    return () => {
      chartRef.current?.destroy();
      chartRef.current = null;
    };
  }, [series]);

  if (isLoading) {
    return (
      <div style={{ padding: '24px', textAlign: 'center', color: '#605e5c' }}>
        Loading sales data…
      </div>
    );
  }

  if (series.length === 0) {
    return (
      <div style={{ padding: '24px', textAlign: 'center', color: '#605e5c' }}>
        No sales data available.
      </div>
    );
  }

  return (
    <div style={{ padding: '16px', fontFamily: 'Segoe UI, sans-serif' }}>
      <div style={{ marginBottom: '12px', fontSize: '14px', fontWeight: 600, color: '#323130' }}>
        Sales by Region &mdash; Total: ${totalSales.toLocaleString()}
      </div>
      <div style={{ position: 'relative', height: '320px' }}>
        <canvas ref={canvasRef} />
      </div>
    </div>
  );
};
