import { Component, Injector } from '@angular/core';
import { BsDatepickerConfig, BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { NgbActiveModal, NgbModal } from '@ng-bootstrap/ng-bootstrap';
import { NgxSpinnerService } from 'ngx-spinner';
import { DatePipe, isPlatformBrowser } from '@angular/common';
import { Title } from '@angular/platform-browser';
import { common } from '../../../../common';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
import { Api } from '../../../../Core/Providers/Api/api';

@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [SharedModule, BsDatepickerModule],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss'
})
export class Dashboard {

  bsConfig!: Partial<BsDatepickerConfig>;
  maxDate: Date = new Date();
  bsValue: Date = new Date();
  selectedDate: Date = new Date();
  constructor(
    public apiSrvc: Api,
    private ngbmodal: NgbModal,
    private ngbmodalActive: NgbActiveModal,
    private spinner: NgxSpinnerService,
    private datepipe: DatePipe,
    private title: Title,
    private comm: common,
    private toast: ToastService,
    private injector: Injector
  ) {
    this.title.setTitle(this.comm.titleName + '-RO Collision Center');
    const data = {
      title: 'RO Collision Center',
      store: 2,
    };
    this.apiSrvc.SetHeaderData({
      obj: data,
    });
    this.GetCollisionCenterData(this.selectedDate);
    this.bsConfig = {
      dateInputFormat: 'MM/DD/YYYY',
      maxDate: this.maxDate,
      showWeekNumbers: false
    };
  }
  onDateSelected(event: Date) {
    if (event) {
      this.selectedDate = event; // store full date
      console.log('Selected Date:', this.selectedDate);
    }
  }
  applyFilterAndYear() {
    if (!this.selectedDate) {
      
      this.toast.show('Please select a date!.', 'danger', 'Error');
      return;
    }
    this.GetCollisionCenterData(this.selectedDate);
  }
  CollisionData: any = [];

  GetCollisionCenterData(date: any) {
    this.CollisionData = [];
    const DateToday = this.datepipe.transform(
      new Date(date),
      'MM/dd/yyyy'
    );
    const obj = {
      "RoDate": DateToday
    };
    this.spinner.show();

    this.apiSrvc.postmethod(this.comm.routeEndpoint + 'GetROCollisionCenter', obj)
      .subscribe(
        (res: any) => {
          this.spinner.hide();
          if (res.status === 200) {
            this.CollisionData = (res.response || [])
            console.log('CollisionData', this.CollisionData);
          }
        },
        (error) => {
          this.spinner.hide();
          console.error('API Error:', error);
        }
      );
  }

  formatCurrency(value: any): string {
    if (value === null || value === undefined || value === '') return '-';
    const num = Number(value);
    return isNaN(num) ? '-' : `$${num.toLocaleString('en-US', { minimumFractionDigits: 0 })}`;
  }

  formatPercent(value: any): string {
    if (value === null || value === undefined || value === '') return '-';
    const num = Number(value);
    return isNaN(num) ? '-' : `${num.toFixed(1)}%`;
  }

}
