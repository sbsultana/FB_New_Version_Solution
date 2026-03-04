import {
  Component,
  ElementRef,
  Injector,
  Input,
  OnInit,
  SimpleChanges,
  ViewChild,
  HostListener
} from '@angular/core';
import { ActivatedRoute, Router } from '@angular/router';
import { NgxSpinnerService } from 'ngx-spinner';
import { NgbActiveModal, NgbModal } from '@ng-bootstrap/ng-bootstrap';
import { Title } from '@angular/platform-browser';
import { CommonModule, CurrencyPipe, DatePipe, NgStyle } from '@angular/common';
import * as FileSaver from 'file-saver';
import { Workbook } from 'exceljs';
import { jsPDF } from 'jspdf';
import html2canvas from 'html2canvas';
import { Subscription } from 'rxjs';
import { Api } from '../../../../Core/Providers/Api/api';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { BsDatepickerConfig, BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { environment } from '../../../../../environments/environment';
import { common } from '../../../../common';
import { Stores } from '../../../../CommonFilters/stores/stores';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
import { FormsModule } from '@angular/forms';
const EXCEL_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
const EXCEL_EXTENSION = '.xlsx';

@Component({
  selector: 'app-dashboard',
  imports: [CommonModule, SharedModule, BsDatepickerModule, FormsModule],
  standalone: true,
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
})
export class Dashboard {

  LendersData: any;
  filteredData: any[] = [];
  selectedstrid = 0;
  selectedDate: Date = new Date();
  currentMonth!: Date;
  Month: any = '';
  groups: any = 1;
  StoreName: any;
  selectedstorevalues: any = [];
  gridvisibility: any;
  bsRangeValue!: Date[];
  groupsArray: any = [];
  storename: any = ''
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  groupId: any = 0;
  storeIds: any = '0';
  stores: any = []
  activePopover: number = -1;
  storesFilterData: any = {
    'groupsArray': this.groupsArray, 'groupId': this.groupId, 'storesArray': this.stores, 'storeids': '1', 'type': 'M', 'others': 'N',
    'groupName': this.groupName, 'storename': this.storename, storecount: null, 'storedisplayname': this.storedisplayname
  };
  constructor(
    public apiSrvc: Api,
    private spinner: NgxSpinnerService,
    private ngbmodal: NgbModal,
    private ngbmodalActive: NgbActiveModal,
    private toast: ToastService,
    private title: Title,
    private datepipe: DatePipe,
    private comm: common
  ) {
    const data = {
      title: 'Lenders Mapping',
      path1: '',
      path2: '',
      path3: '',
      stores: this.storeIds,
      groups: this.groups,
      count: 0
    };
    this.apiSrvc.SetHeaderData({
      obj: data,
    });
    this.GetLenders(this.selectedstrid);
  }
  NoData: any = false;
  searchText: any = '';
  GetLenders(strid: any) {
    this.LendersData = [];
    this.NoData = false;
    this.spinner.show();
    const obj = {
      store_id: strid,
      accountname: this.searchText,
      Lendertype: '0',
      LenderCategory: '',
      STATUS: 'A',
      UserID: 0,
    };
    const curl = environment.apiUrl + this.comm.routeEndpoint + 'GetLenderTab';

    this.apiSrvc.postmethod(this.comm.routeEndpoint + 'GetLenderTab', obj).subscribe(
      (res) => {
        const currentTitle = document.title;
        this.apiSrvc.logSaving(curl, {}, '', res.message, currentTitle);
        if (res.status == 200) {
          this.spinner.hide();
          this.LendersData = res.response;
          this.filteredData = this.LendersData;
        } else {
          this.toast.show(res.status, 'danger', 'Error');
          this.spinner.hide();
          this.NoData = true;
        }
      },
      (error) => {
        this.toast.show('502 Bad Gate Way Error', 'danger', 'Error');
        this.spinner.hide();
        this.NoData = true;
      }
    );
  }
  Scrollpercent: any = 0;
  scrollCurrentposition: any = 0
  @ViewChild('scrollcent') scrollcent!: ElementRef;

  updateVerticalScroll(event: any): void {
    this.scrollCurrentposition = event.target.scrollTop
    const scrollDemo = document.querySelector('#scrollcent') as HTMLElement;
    this.Scrollpercent = Math.round(
      (event.target.scrollTop /
        (event.target.scrollHeight - scrollDemo.clientHeight)) *
      100
    );
  }
  onSearch() {
    const value = this.searchText.toLowerCase();

    this.filteredData = this.LendersData.filter((item: any) =>
      (item.DMS_Lender_Text || '').toLowerCase().includes(value) ||
      (item.LenderName || '').toLowerCase().includes(value) ||
      (item.LenderType || '').toLowerCase().includes(value) ||
      (item.Category || '').toLowerCase().includes(value)
    );
  }
}
