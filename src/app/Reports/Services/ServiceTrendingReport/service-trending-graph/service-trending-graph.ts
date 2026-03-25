import { Component, OnInit, ViewChild, Input, Renderer2 } from '@angular/core';
import { Chart } from 'chart.js';
import { NgbActiveModal, NgbModal } from '@ng-bootstrap/ng-bootstrap';
import { getRelativePosition } from 'chart.js/helpers';
import { CommonModule } from '@angular/common';

import numeral from 'numeral';

@Component({
  selector: 'app-service-trending-graph',
  imports: [CommonModule],
  templateUrl: './service-trending-graph.html',
  styleUrl: './service-trending-graph.scss',
})
export class ServiceTrendingGraph {
  @Input() STRgraphdetails: any;
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
    console.log(this.STRgraphdetails.Object);
    this.itemdata = { ...this.STRgraphdetails.Object };
    const Val = 'Store_Name';
    const Val1 = 'ASD_AS_ID';
    const Val2 = 'ast_ltype';
    const Val3 = 'paytype';
    const Val4 = 'variable';
    const Val5 = 'seq';

    if (this.itemdata.hasOwnProperty(Val)) {
      delete this.itemdata[Val];
    }
    if (this.itemdata.hasOwnProperty(Val1)) {
      delete this.itemdata[Val1];
    }
    if (this.itemdata.hasOwnProperty(Val2)) {
      delete this.itemdata[Val2];
    }
    if (this.itemdata.hasOwnProperty(Val3)) {
      delete this.itemdata[Val3];
    }
    if (this.itemdata.hasOwnProperty(Val4)) {
      delete this.itemdata[Val4];
    }
    if (this.itemdata.hasOwnProperty(Val5)) {
      delete this.itemdata[Val5];
    }
    console.log('Graph Object', this.itemdata);

    this.monthvalues = Object.values(this.itemdata);
    this.monthkeys = this.STRgraphdetails.DATES;
    console.log(this.monthkeys, "monthkeys");
    console.log(this.monthvalues, 'monthvalues');
    const formattedDates = this.monthkeys.map((date: any) =>
      this.formatDate(date)
    );
    this.monthkeys = [];
    formattedDates.forEach((date: any) => this.monthkeys.push(date));
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
    console.log(this.monthvalues);
    const ETPaceValue = this.monthvalues.map((_arrayElement: any) =>
      Object.assign({}, _arrayElement)
    );
    var d1 = [];

    for (var i = 0; i < this.monthvalues.length; i++) {
      console.log(this.monthvalues[i], this.monthvalues);
      console.log(i, this.monthvalues[i]);
      d1.push([i, this.monthvalues[i]]);
    }
    let d2 = this.lineFit(d1);
    ETPaceValue.pop();
    ETPaceValue.push(this.STRgraphdetails.Object.PACE);
    for (let i = 0; i < ETPaceValue.length; i++) {
      this.pace.push(this.STRgraphdetails.Object.PACE);
    }
    let myChart = new Chart(this.ctx, {
      type: 'line',
      data: {
        labels: this.monthkeys.slice(1),
        datasets: [
          {
            label: this.STRgraphdetails.NAME,
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
            data: this.pace,
            label: 'Pace',
            type: 'line',
            fill: false,
            borderColor: '#ffab00',
            borderWidth: 2,
            borderDash: [6, 6],
            pointRadius: 0,
          },
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
                const value = context.raw;
                return this.ValueFormat(value);
              }
            }
          }
        }
      }
    });
  }
  // private ValueFormat(value: number) {
  //   if (this.STRgraphdetails.ValueFormat == 'Number') {
  //     const formattedValue = numeral(value).format('0,0');
  //     return formattedValue;
  //   } else if (this.STRgraphdetails.ValueFormat == 'Percentage') {
  //     const formattedValue = numeral(value * 1).format('0') + '%';
  //     return formattedValue;
  //   } else if (this.STRgraphdetails.ValueFormat == 'Currancy') {
  //     return numeral(value).format('($0,0)').replace('(', '-').replace(')', '');
  //   }
  // }
  private ValueFormat(value: number) {
    if (this.STRgraphdetails.ValueFormat == 'Number') {
      return numeral(value).format('0,0');
    }
    else if (this.STRgraphdetails.ValueFormat == 'Percentage') {
      return numeral(value).format('0') + '%';
    }
    else if (this.STRgraphdetails.ValueFormat == 'Currancy') {
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

