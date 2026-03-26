import { Component, OnInit, ViewChild, Input, Renderer2 } from '@angular/core';
import { Chart } from 'chart.js';
import { NgbActiveModal, NgbModal } from '@ng-bootstrap/ng-bootstrap';
import { getRelativePosition } from 'chart.js/helpers';
import { CommonModule } from '@angular/common';

import numeral from 'numeral';

@Component({
  selector: 'app-parts-trending-graph',
  imports: [CommonModule],
  templateUrl: './parts-trending-graph.html',
  styleUrl: './parts-trending-graph.scss',
})
export class PartsTrendingGraph {
  @Input() PTRgraphdetails: any;
  canvas: any;
  ctx: any;
  @ViewChild('mychart', { static: true }) mychart: any;
  monthvalues: any = [];
  monthkeys: any = [];
  constructor(
    private ngbmodalActive: NgbActiveModal,
    private renderer: Renderer2
  ) {
    this.renderer.listen('window', 'click', (e: Event) => {
      const TagName = e.target as HTMLButtonElement;
      console.log(TagName.className);
      if (TagName.className === 'd-block modal fade show modal-static') {
        this.close();
      }
    });
  }
  itemdata: any;

  ngOnInit(): void {
    console.log(this.PTRgraphdetails.Object);

    const rawData = { ...this.PTRgraphdetails.Object };
    const removeKeys = ['Store_Name', 'AP_StoreID', 'paytype', 'variable'];

    this.itemdata = Object.fromEntries(
      Object.entries(rawData).filter(([key]) => !removeKeys.includes(key))
    );
    this.monthkeys = this.PTRgraphdetails.DATES || Object.keys(this.itemdata);
    this.monthvalues = this.monthkeys.map((key: any) => this.itemdata[key] ?? 0);
    this.monthkeys = this.monthkeys.map((date: any) => this.formatDate(date));

    console.log(this.monthkeys, 'monthkeys');
    console.log(this.monthvalues, 'monthvalues');
  }
  formatDate(dateStr: string): string {
    const [year, month] = dateStr.split(' ');
    const date = new Date(`${month} 1, ${year}`);
    const monthAbbreviation = date.toLocaleString('default', {
      month: 'short',
    });
    const fullYear = date.getFullYear();
    return `${monthAbbreviation} ${fullYear}`;
  }
  pace: any = [];
  ngAfterViewInit() {
    this.canvas = this.mychart.nativeElement;
    this.ctx = this.canvas.getContext('2d');

    // ✅ Pace line (straight line)
    this.pace = this.monthvalues.map(() => this.PTRgraphdetails.Object.PACE);

    new Chart(this.ctx, {
      type: 'line',
      data: {
        labels: this.monthkeys, // ❌ removed slice(1)
        datasets: [
          {
            label: this.PTRgraphdetails.NAME,
            borderColor: '#00c2ff',
            backgroundColor: 'rgba(0, 194, 255, 0.15)',
            fill: true,
            data: this.monthvalues,
            tension: 0.35,
            pointBackgroundColor: '#ffffff',
            pointBorderColor: '#00c2ff',
            pointRadius: 4,
            borderWidth: 2,
          },
          {
            label: 'Pace',
            data: this.pace,
            borderColor: '#ffab00',
            borderWidth: 2,
            borderDash: [6, 6],
            pointRadius: 0,
            fill: false,
          }
        ],
      },
      options: {
        responsive: true,

        plugins: {
          legend: {
            display: true,
            position: 'bottom',
            labels: {
              usePointStyle: true,
              color: 'white'
            }
          },

          tooltip: {
            callbacks: {
              label: (context: any) => {
                return this.ValueFormat(context.raw);
              }
            }
          }
        },

        scales: {
          x: {
            ticks: {
              color: 'white'
            },
            grid: {
              display: false
            }
          },
          y: {
            ticks: {
              color: 'white',
              callback: (value: any) => this.ValueFormat(value)
            },
            grid: {
              color: 'rgba(255,255,255,0.1)'
            }
          }
        }
      }
    });
  }
  // private ValueFormat(value: number) {
  //   if (this.PTRgraphdetails.ValueFormat == 'Number') {
  //     const formattedValue = numeral(value).format('0,0');
  //     return formattedValue;
  //   } else if (this.PTRgraphdetails.ValueFormat == 'Percentage') {
  //     const formattedValue = numeral(value * 1).format('0') + '%';
  //     return formattedValue;
  //   } else if (this.PTRgraphdetails.ValueFormat == 'Currancy') {
  //     return numeral(value).format('($0,0)').replace('(', '-').replace(')', '');
  //   }
  // }
  private ValueFormat(value: number) {
    if (this.PTRgraphdetails.ValueFormat == 'Number') {
      return numeral(value).format('0,0');
    }
    else if (this.PTRgraphdetails.ValueFormat == 'Percentage') {
      return numeral(value).format('0') + '%';
    }
    else if (this.PTRgraphdetails.ValueFormat == 'Currancy') {
      return numeral(value)
        .format('($0,0)')
        .replace('(', '-')
        .replace(')', '');
    }

    // ✅ Default return (important)
    return value?.toString() ?? '';
  }
  private formatPercentage(value: number): string {
    const formattedValue = numeral(value * 1).format('0') + '%';
    return formattedValue;
  }
  private formatNumber(value: number): string {
    const formattedValue = numeral(value).format('0,0');
    return formattedValue;
  }
  rV: any;
  lineFit(points: any) {
    let sI: any = this.slopeAndIntercept(points);
    //con.log(sI)
    if (sI) {
      var N = points.length;
      this.rV = [];
      //  this.rV.push([points[0][0], sI['slope'] * points[0][0] + sI['intercept']]);
      //  this. rV.push([points[N-1][0], sI['slope'] * points[N-1][0] + sI['intercept']]);
      for (let i = 0; i < points.length; i++) {
        this.rV.push(sI['slope'] * points[i][0] + sI['intercept']);
      }
      //con.log( this.rV)
      return this.rV;
    }
    return [];
  }
  slopeAndIntercept(points: any) {
    var rV: any = {},
      N = points.length,
      sumX = 0,
      sumY = 0,
      sumXx = 0,
      sumYy = 0,
      sumXy = 0;

    if (N < 2) {
      return rV;
    }

    for (var i = 0; i < N; i++) {
      var x = points[i][0],
        y = points[i][1];
      sumX += x;
      sumY += y;
      sumXx += x * x;
      sumYy += y * y;
      sumXy += x * y;
    }
    rV['slope'] = (N * sumXy - sumX * sumY) / (N * sumXx - sumX * sumX);
    rV['intercept'] = (sumY - rV['slope'] * sumX) / N;
    rV['rSquared'] = Math.abs(
      (rV['slope'] * (sumXy - (sumX * sumY) / N)) / (sumYy - (sumY * sumY) / N)
    );
    // //con.log(rV)
    return rV;
  }

  close() {
    this.ngbmodalActive.close();
  }
}

