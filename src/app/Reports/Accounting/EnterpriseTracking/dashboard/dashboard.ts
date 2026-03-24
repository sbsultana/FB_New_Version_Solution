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
import { EnterprisetrackingDetails } from '../enterprisetracking-details/enterprisetracking-details';
import { Stores } from '../../../../CommonFilters/stores/stores';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
const EXCEL_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
const EXCEL_EXTENSION = '.xlsx';



@Component({
  selector: 'app-dashboard',
  imports: [CommonModule, SharedModule, BsDatepickerModule, Stores],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss'
})
export class Dashboard implements OnInit {
  reportOpenSub!: Subscription;
  reportGetting!: Subscription;
  Pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;
  date: any;
  StoreValues: any = [];
  Month: any = '';
  groups: any = 1;
  StoreName: any;
  SpcClick: any;
  block: any = '';
  FromDate: any;
  ToDate: any;
  EnterpriseTrackingData: any = [];
  NodataFound!: boolean;
  NoData!: boolean;
  header: any = [
    {
      type: 'Bar',
      StoreValues: this.StoreValues,
      Month: this.Month,
      groups: this.groups,
    },
  ];
  popup: any = [{ type: 'Popup' }];
  selectedDate: Date = new Date();
  currentMonth!: Date;
  // StoreValues: any = '0';
  // popup: any = [{ type: 'Popup' }];
  // groups: any = 1;
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
    private route: ActivatedRoute,
    private router: Router,
    private spinner: NgxSpinnerService,
    private ngbmodal: NgbModal,
    private ngbmodalActive: NgbActiveModal,
    private title: Title,
    // private excelservice: ExcelService,
    private datepipe: DatePipe,
    private comm: common,
    private injector: Injector,
    public shared: Sharedservice,
    private toast: ToastService
  ) {
    const lastMonth = new Date();
    let today = new Date();
    if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
      // this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.ustores.split(',')
      this.groupId = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Preferences
      this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids.split(',')
    }
    console.log();
    if (this.shared.common.groupsandstores.length > 0) {
      this.groupsArray = this.shared.common.groupsandstores.filter((val: any) => val.sg_id != this.shared.common.reconID);
      this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
      this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_Name : this.groupName = ''
      this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
      this.getStoresandGroupsValues()
    }
    if (today.getDate() < 5) {
      this.date = new Date(lastMonth.setMonth(lastMonth.getMonth() - 1));
    } else {
      this.date = new Date(lastMonth.setMonth(lastMonth.getMonth()));
    }
    this.Month =
      this.date.toString().substr(8, 2) +
      '-' +
      this.date.toString().substr(4, 3) +
      '-' +
      this.date.toString().substr(11, 4);
    this.title.setTitle(this.comm.titleName + '-Enterprise Tracking');
    const data = {
      title: 'Enterprise Tracking',
      path1: '',
      path2: '',
      path3: '',
      Month: this.date,
      stores: this.StoreValues.toString(),
      store: this.StoreValues,
      groups: this.groups,
      count: 0,
    };
    this.apiSrvc.SetHeaderData({
      obj: data,
    });
    this.header = [
      {
        type: 'Bar',
        StoreValues: this.StoreValues,
        Month: this.date,
        groups: this.groups,
      },
    ];
    this.selectedDate = this.date
    this.currentMonth = this.selectedDate;
    this.GetEnterpriseTracking(this.currentMonth);
  }

  bsConfig: Partial<BsDatepickerConfig> = {
    dateInputFormat: 'MMMM/YYYY',
    minMode: 'month',
    maxDate: new Date()
  };

  applyDateChange() {
    if (!this.storeIds || this.storeIds.length === 0) {
    
      this.toast.show(
        'Please Select Atleast One Store',
        'warning',
        'Warning'
      );
      return;
    }
    else {
      this.currentMonth = this.selectedDate;
      this.GetEnterpriseTracking(this.currentMonth);
    }
  }

  Scrollpercent: any = 0;
  scrollCurrentposition: any = 0;
  @ViewChild('scrollcent') scrollcent!: ElementRef;

  updateVerticalScroll(event: any): void {
    this.scrollCurrentposition = event.target.scrollTop;
    const scrollDemo = document.querySelector('#scrollcent') as HTMLElement;
    this.Scrollpercent = Math.round(
      (event.target.scrollTop /
        (event.target.scrollHeight - scrollDemo.clientHeight)) *
      100
    );
  }

  storeName: any;
  LastMonth: any = '';
  GetEnterpriseTracking(date: any) {
    this.spinner.show();

    const Month = this.datepipe.transform(new Date(date), 'yyyy-MM-dd');
    const lastMonthDate = new Date(date);
    lastMonthDate.setFullYear(lastMonthDate.getFullYear() - 1);
    this.LastMonth = lastMonthDate;
    let obj = {
      AS_IDS: this.storeIds.toString(),
      DATE: Month,
    };
    console.log(obj);
    const curl = environment.apiUrl + this.comm.routeEndpoint + 'GetEnterpriseTrackingNetprofit';
    this.apiSrvc
      .postmethod(this.comm.routeEndpoint + 'GetEnterpriseTrackingNetprofit', obj)
      .subscribe(
        (res) => {
          const currentTitle = document.title;
          this.apiSrvc.logSaving(curl, {}, '', res.message, currentTitle);
          if (res.status == 200) {
            this.EnterpriseTrackingData = res.response;
            this.EnterpriseTrackingData = res.response.reduce(
              (r: any, { STORE, ...rest }: any) => {
                if (!r.some((o: any) => o.STORE == STORE)) {
                  r.push({
                    STORE,
                    ...rest,
                    sub: res.response.filter((v: any) => v.STORE == STORE),
                  });
                }
                return r;
              },
              []
            );
            console.log(
              'Enterprise Tracking Data',
              this.EnterpriseTrackingData
            );
            this.spinner.hide();
            this.NodataFound = true;
            if (this.EnterpriseTrackingData.length > 0) {
              this.NoData = false;
            } else {
              this.NoData = true;
            }
          } else {
            this.spinner.hide();
           
            this.toast.show('Invalid Details.', 'danger', 'Error');
          }
        },
        (error) => {
          console.log(error);
        }
      );
    // });
  }

  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    }
    return false;
  }

  openMenu(Object: any, LatestDate: any, Type: any) {
    console.log(Object);
    const ETdetails = this.ngbmodal.open(EnterprisetrackingDetails, {
      size: 'xl',
      backdrop: 'static',
      injector: Injector.create({
        providers: [
          { provide: CurrencyPipe, useClass: CurrencyPipe }
        ],
        parent: this.injector
      })
    });
    ETdetails.componentInstance.ETdetails = {
      TYPE: Type,
      DEPARTMENT: Object.DEPT,
      COMPANYID: Object.STOREID,
      MONTH: LatestDate,
      STORE: Object.STORE,
      Group: this.groups,
    };
  }

  ngOnInit(): void {
    // this.GetSalesReconciliationData();

    // var curl =
    //   'https://fbxtractapi.axelautomotive.com/api/fredbeans/GetSalesPersonCommCalc';
    // this.apiSrvc.logSaving(curl, {}, '', 'Success', 'Enterprise Tracking');
    this.getStores();
  }
  getStoresandGroupsValues() {
    this.storesFilterData.groupsArray = this.groupsArray;
    this.storesFilterData.groupId = this.groupId;
    this.storesFilterData.storesArray = this.stores;
    this.storesFilterData.storeids = this.storeIds;
    this.storesFilterData.groupName = this.groupName;
    this.storesFilterData.storename = this.storename;
    this.storesFilterData.storecount = this.storecount;
    this.storesFilterData.storedisplayname = this.storedisplayname;

    this.storesFilterData = {
      groupsArray: this.groupsArray,
      groupId: this.groupId,
      storesArray: this.stores,
      storeids: this.storeIds,
      groupName: this.groupName,
      storename: this.storename,
      storecount: this.storecount,
      storedisplayname: this.storedisplayname,
      'type': 'M', 'others': 'N'
    };

    // this.setHeaderData();
    // this.GetData();

  }
  StoresData(data: any) {
    this.storeIds = data.storeids;
    this.groupId = data.groupId;
    this.storename = data.storename;
    this.groupName = data.groupName;
    this.storecount = data.storecount;
    this.storedisplayname = data.storedisplayname;
  }
  getStores() {
    this.selectedstorevalues = [];
  }
  ngAfterViewInit(): void {
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.shared.common.pageName == 'Enterprise Tracking') {
        if (res.obj.storesData != undefined) {
          this.groupsArray = res.obj.storesData;
          this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
          this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_name : this.groupName = ''
          this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.StoreValues)[0].storename : this.storename = ''
          this.getStoresandGroupsValues()
        }
      }
    })
    this.reportOpenSub = this.apiSrvc.GetReportOpening().subscribe((res) => {
      // console.log(res);
      if (this.reportOpenSub != undefined) {
        if (res.obj.Module == 'Enterprise Tracking') {
          document.getElementById('report')?.click();
        }
      }
    });
    this.reportGetting = this.apiSrvc.GetReports().subscribe((data) => {
      console.log(data);
      if (this.reportGetting != undefined) {
        if (data.obj.Reference == 'Enterprise Tracking') {
          if (data.obj.header == undefined) {
            this.date = data.obj.month;
            this.Month = data.obj.month;
            console.log(data.obj.month);
            this.storeIds = data.obj.storeValues;
            this.StoreName = data.obj.Sname;
            this.groups = data.obj.groups;
          } else {
            if (data.obj.header == 'Yes') {
              this.storeIds = data.obj.storeValues;
            }
          }
          this.GetEnterpriseTracking(this.currentMonth);
          const headerdata = {
            title: 'Enterprise Tracking',
            path1: '',
            path2: '',
            path3: '',
            Month: this.Month,
            stores: this.storeIds,
            groups: this.groups,
          };
          this.apiSrvc.SetHeaderData({
            obj: headerdata,
          });
          this.header = [
            {
              type: 'Bar',
              storeIds: this.storeIds,
              Month: new Date(this.Month),
              groups: this.groups,
            },
          ];
        }
      }
    });
    this.excel = this.apiSrvc.getExportToExcelAllReports().subscribe((res) => {
      if (this.excel != undefined) {
        this.SpcClick = res.obj.state;
        if (res.obj.title == 'Enterprise Tracking') {
          if (res.obj.state == true) {
            this.exportAsXLSX();
          }
        }
      }
    });
    this.email = this.apiSrvc
      .getExportToEmailPDFAllReports()
      .subscribe((res) => {
        if (this.email != undefined) {
          if (res.obj.title == 'Enterprise Tracking') {
            if (res.obj.stateEmailPdf == true) {
              this.sendEmailData(res.obj.Email, res.obj.notes, res.obj.from);
            }
          }
        }
      });

    this.Pdf = this.apiSrvc.getExportToPDFAllReports().subscribe((res) => {
      if (this.Pdf != undefined) {
        if (res.obj.title == 'Enterprise Tracking') {
          if (res.obj.statePDF == true) {
            this.generatePDF();
          }
        }
      }
    });
    this.print = this.apiSrvc.getExportToPrintAllReports().subscribe((res) => {
      if (this.print != undefined) {
        if (res.obj.title == 'Enterprise Tracking') {
          if (res.obj.statePrint == true) {
            this.GetPrintData();
          }
        }
      }
    });
  }


  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent) {
    const clickedInside = (event.target as HTMLElement).closest('.dropdown-toggle, .reportstores-card, .timeframe');
    if (!clickedInside) {
      this.activePopover = -1;
    }
  }

  togglePopover(popoverIndex: number) {
    this.activePopover = this.activePopover === popoverIndex ? -1 : popoverIndex;
  }

  reportOpen(temp: any) {
    this.ngbmodalActive = this.ngbmodal.open(temp, {
      size: 'xl',
      backdrop: 'static',
    });
  }
getSelectedStoreNames(): string {
    if (!this.storeIds || this.storeIds.length === 0) return '';

    const ids = this.storeIds.toString().split(',');

    const selectedStores = this.stores.filter((s: any) =>
      ids.includes(s.ID.toString())
    );

    return selectedStores.map((s: any) => s.storename).join(', ');
  }
  getReportFilters(): { title: string; filters: any[] } {
    return {
      title: 'Expense Trend',
      filters: [
        {
          label: 'Store',
          value: this.getSelectedStoreNames() || 'All Stores'
        },
        {
          label: 'Group',
          value: this.groupName || ''
        },
        {
          label: 'Month',
          value: this.datepipe.transform(this.currentMonth, 'MMMM yyyy')
        }
      ]
    };
  }
  addExcelFiltersSection(worksheet: any): number {
    let rowCount = 0;

    const report = this.getReportFilters();

    /*  TITLE (LEFT ALIGNED) */
    const titleRow = worksheet.addRow([report.title]);
    titleRow.font = { bold: true, size: 14 };
    worksheet.mergeCells(`A${rowCount + 1}:G${rowCount + 1}`);
    titleRow.alignment = { horizontal: 'left', vertical: 'middle' };
    rowCount++;

    /* FILTERS */
    report.filters.forEach((filter: any) => {
      const row = worksheet.addRow([`${filter.label}:`, filter.value]);
      row.getCell(1).font = { bold: true };
      rowCount++;
    });

    /* SPACE */
    worksheet.addRow([]);
    rowCount++;

    return rowCount;
  }

  ExcelStoreNames: any = [];
  exportAsXLSX(): void {


    let storeValue = '';
    const selectedStoreIds: string[] = this.storeIds && this.storeIds.length
      ? this.storeIds.map((id: any) => id.toString()) : [];
    const allStores: any[] = Array.isArray(this.stores) ? this.stores : [];
    storeValue = allStores.filter((s: any) => selectedStoreIds.includes(s.ID.toString()))
      .map((s: any) => s.storename.trim())
      .filter(Boolean)
      .join(', ');

    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet('Enterprise Tracking');
    worksheet.views = [
      {
        state: 'frozen',
        ySplit: 13, // Number of rows to freeze (2 means the first two rows are frozen)
        topLeftCell: 'A14', // Specify the cell to start freezing from (in this case, the third row)
        showGridLines: false,
      },
    ];
    worksheet.addRow('');
    const titleRow = worksheet.addRow(['Enterprise Tracking']);
    titleRow.eachCell((cell, number) => {
      cell.alignment = {
        indent: 1,
        vertical: 'middle',
        horizontal: 'left',
      };
    });
    titleRow.font = { name: 'Arial', family: 4, size: 12, bold: true };
    titleRow.worksheet.mergeCells('A2', 'D2');

    worksheet.addRow('');
    const DateToday = this.datepipe.transform(
      new Date(),
      'MM/dd/yyyy h:mm:ss a'
    );
    worksheet.addRow([DateToday]).font = {
      name: 'Arial',
      family: 4,
      size: 9,
    };
    const EnterpriseTrackingData = this.EnterpriseTrackingData.map(
      (_arrayElement: any) => Object.assign({}, _arrayElement)
    );
    const LastYear = this.datepipe.transform(
      new Date(this.LastMonth),
      'MMM yyyy'
    );
    const Header = [
      'Stores',
      'Units',
      'Gross',
      'Pace',
      LastYear,
      '% Diff',
      'Variable',
      'Personnel/Selling',
      'Semi-Fixed/Operating',
      'Fixed/Overhead',
      'Total',
      'Op. Net',
      LastYear,
      'Difference',
      'Net Adjust',
      'Profit/Loss',
    ];
    // for (let i = 0; i < this.lastFourMonthsArray.length; i++) {
    //  const formattedDate = this.datepipe.transform(this.lastFourMonthsArray[i], 'MMMM yyyy');
    //   Header.push(formattedDate!);
    // }
    const ReportFilter = worksheet.addRow(['Report Filters :']);
    ReportFilter.font = { name: 'Arial', family: 4, size: 10, bold: true };

    const SummaryType = worksheet.addRow(['Summary Type :']);
    const summarytype = worksheet.getCell('B6');
    summarytype.value = 'Store Summary';
    summarytype.font = { name: 'Arial', family: 4, size: 9 };
    summarytype.alignment = { vertical: 'top', horizontal: 'left' };
    SummaryType.getCell(1).font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
    };
    const DateMonth = worksheet.addRow(['Month :']);
    const datemonth = worksheet.getCell('B7');
    datemonth.value = this.currentMonth;
    datemonth.font = { name: 'Arial', family: 4, size: 9 };
    datemonth.alignment = { vertical: 'top', horizontal: 'left' };
    DateMonth.getCell(1).font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
    };
    const Groups = worksheet.getCell('A8');
    Groups.value = 'Group :';
    Groups.font = { name: 'Arial', family: 4, size: 9, bold: true };
    const groups = worksheet.getCell('B8');
    groups.value = 'Silvertip';
    groups.font = { name: 'Arial', family: 4, size: 9 };
    groups.alignment = {
      vertical: 'middle',
      horizontal: 'left',
      wrapText: true,
    };
    worksheet.mergeCells('B9', 'K11');
    const Stores = worksheet.getCell('A9');
    Stores.value = 'Store :';
    const stores = worksheet.getCell('B9');
    stores.value = storeValue;
    stores.font = { name: 'Arial', family: 4, size: 9 };
    stores.alignment = {
      vertical: 'top',
      horizontal: 'left',
      wrapText: true,
    };
    Stores.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
    };
    worksheet.addRow('');

    const row = worksheet.getRow(12);
    row.height = 20;

    // const FromDate = this.datepipe.transform( new Date(this.FromDate),'MMM dd');
    // const ToDate = this.datepipe.transform( new Date(this.ToDate),'MMM dd');
    let DateRange = worksheet.getCell('A12');
    DateRange.value = this.selectedDate;
    DateRange.alignment = { vertical: 'middle', horizontal: 'center' };
    DateRange.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
      color: { argb: 'FFFFFF' },
    };
    DateRange.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '2a91f0' },
      bgColor: { argb: 'FF0000FF' },
    };
    DateRange.border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' },
    };
    worksheet.mergeCells('B12');
    let Uspace = worksheet.getCell('B12');
    Uspace.value = 'Retail';
    Uspace.alignment = { vertical: 'middle', horizontal: 'center' };
    Uspace.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
      color: { argb: 'FFFFFF' },
    };
    Uspace.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '2a91f0' },
      bgColor: { argb: 'FF0000FF' },
    };
    Uspace.border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' },
    };

    worksheet.mergeCells('C12', 'F12');
    let FrontGross = worksheet.getCell('C12');
    FrontGross.value = 'Gross';
    FrontGross.alignment = { vertical: 'middle', horizontal: 'center' };
    FrontGross.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
      color: { argb: 'FFFFFF' },
    };
    FrontGross.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '2a91f0' },
      bgColor: { argb: 'FF0000FF' },
    };
    FrontGross.border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' },
    };

    worksheet.mergeCells('G12', 'K12');
    let BackGross = worksheet.getCell('G12');
    BackGross.value = 'Estimated Expenses';
    BackGross.alignment = { vertical: 'middle', horizontal: 'center' };
    BackGross.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
      color: { argb: 'FFFFFF' },
    };
    BackGross.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '2a91f0' },
      bgColor: { argb: 'FF0000FF' },
    };
    BackGross.border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' },
    };

    worksheet.mergeCells('L12', 'N12');
    let Commission = worksheet.getCell('L12');
    Commission.value = 'Operating Net';
    Commission.alignment = { vertical: 'middle', horizontal: 'center' };
    Commission.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
      color: { argb: 'FFFFFF' },
    };
    Commission.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '2a91f0' },
      bgColor: { argb: 'FF0000FF' },
    };
    Commission.border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' },
    };

    worksheet.mergeCells('O12', 'P12');
    let Bonuses = worksheet.getCell('O12');
    Bonuses.value = 'Net Profit';
    Bonuses.alignment = { vertical: 'middle', horizontal: 'center' };
    Bonuses.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
      color: { argb: 'FFFFFF' },
    };
    Bonuses.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '2a91f0' },
      bgColor: { argb: 'FF0000FF' },
    };
    Bonuses.border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' },
    };

    // worksheet.mergeCells("L12", "N12");
    // let TotalPay = worksheet.getCell("L12");
    // TotalPay.value = "Total Pay";
    // TotalPay.alignment = { vertical: 'middle', horizontal: 'center' };
    // TotalPay.font = { name: 'Arial',family: 4,size: 9,bold: true, color: { argb: 'FFFFFF' }, }
    // TotalPay.fill = {type: 'pattern',pattern: 'solid',fgColor: { argb: '2a91f0' },bgColor: { argb: 'FF0000FF' }};
    // TotalPay.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } }

    const headerRow = worksheet.addRow(Header);
    console.log(headerRow);
    headerRow.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
      color: { argb: 'FFFFFF' },
    };
    headerRow.alignment = {
      indent: 1,
      vertical: 'middle',
      horizontal: 'center',
    };
    headerRow.height = 20;
    headerRow.eachCell((cell, number) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '788494' },
        bgColor: { argb: 'FF0000FF' },
      };
      cell.border = { right: { style: 'thin' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });
    // worksheet.eachRow((row, rowNumber) => {
    //   if (rowNumber >= 14) {
    //     row.eachCell((cell, colNumber) => {
    //       if (colNumber !== 1) {  
    //         cell.alignment = {
    //           indent: 1,
    //           vertical: 'middle',
    //           horizontal: 'right'
    //         };
    //       }
    //     });
    //   }
    // });

    for (const d of EnterpriseTrackingData) {

      // =========================
      // MAIN ROW (STORE LEVEL)
      // =========================

      const Data1 = worksheet.addRow([
        d.STORE ?? '-',
        d.RETAILUNITS_UNITS ?? '-',
        d.GROSS_PACE ?? '-',
        d.PACE ?? '-',
        d.GROSS_YOY ?? '-',
        d.GROSS_DIFFPERCENT ?? '-',
        d.EXP_VARIABLE ?? '-',
        d.EXP_PERSONNEL ?? '-',
        d.EXP_SEMIFIXED ?? '-',
        d.EXP_FIXED ?? '-',
        d.EXP_TOTAL ?? '-',
        d.OPERATINGNET_OPNET ?? '-',
        d.OPERATINGNET_YOY ?? '-',
        d.OPERATINGNET_DIFF ?? '-',
        d.NETPROFIT_NETADJUST ?? '-',
        d.NETPROFIT ?? '-'
      ]);

      Data1.font = { name: 'Arial', size: 9 };

      Data1.eachCell((cell, colNumber) => {

        // Border
        cell.border = { right: { style: 'dotted' } };

        // Alignment
        if (colNumber === 1) {
          cell.alignment = {
            indent: 2,
            vertical: 'middle',
            horizontal: 'left'
          };
        } else {
          cell.alignment = {
            indent: 1,
            vertical: 'middle',
            horizontal: 'right'
          };
        }

        // Number formatting
        if (colNumber === 2 || colNumber === 6) {
          cell.numFmt = '#,##0';
        } else {
          cell.numFmt = '$#,##0.00';
        }
      });

      // Zebra Row Background
      if (Data1.number % 2) {
        Data1.eachCell(cell => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFE5E5E5' }
          };
        });
      }

      // =========================
      // SUB ROWS (DEPARTMENT LEVEL)
      // =========================

      if (d.sub?.length) {

        for (let i = 1; i < d.sub.length; i++) {

          const d1 = d.sub[i];

          const Data2 = worksheet.addRow([
            d1.DEPT ?? '-',
            d1.RETAILUNITS_UNITS ?? '-',
            d1.GROSS_PACE ?? '-',
            d1.PACE ?? '-',
            d1.GROSS_YOY ?? '-',
            d1.GROSS_DIFFPERCENT ?? '-',
            d1.EXP_VARIABLE ?? '-',
            d1.EXP_PERSONNEL ?? '-',
            d1.EXP_SEMIFIXED ?? '-',
            d1.EXP_FIXED ?? '-',
            d1.EXP_TOTAL ?? '-',
            d1.OPERATINGNET_OPNET ?? '-',
            d1.OPERATINGNET_YOY ?? '-',
            d1.OPERATINGNET_DIFF ?? '-',
            d1.NETPROFIT_NETADJUST ?? '-',
            d1.NETPROFIT ?? '-'
          ]);

          // Grouping
          Data2.outlineLevel = 1;

          Data2.font = { name: 'Arial', size: 9 };

          Data2.eachCell((cell, colNumber) => {

            // Border
            cell.border = { right: { style: 'dotted' } };

            // Alignment
            if (colNumber === 1) {
              cell.alignment = {
                indent: 3,
                vertical: 'middle',
                horizontal: 'left'
              };
            } else {
              cell.alignment = {
                indent: 1,
                vertical: 'middle',
                horizontal: 'right'
              };
            }

            // Number formatting
            if (colNumber === 2 || colNumber === 6) {
              cell.numFmt = '#,##0';
            } else {
              cell.numFmt = '$#,##0';
            }
          });

          // Zebra Row Background
          if (Data2.number % 2) {
            Data2.eachCell(cell => {
              cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFE5E5E5' }
              };
            });
          }
        }
      }
    }


    worksheet.eachRow((row, rowIndex) => {
      row.eachCell((cell, colIndex) => {
        if (rowIndex > 1 && rowIndex < 11) {
          if (colIndex === 1) {
            cell.alignment = {
              horizontal: 'left',
              vertical: 'middle',
              indent: 1,
            };
          }
        }
      });
    });
    worksheet.getColumn(1).width = 30;
    // worksheet.getColumn(1).alignment = {
    //   indent: 1,
    //   vertical: 'middle',
    //   horizontal: 'left',
    // };
    worksheet.getColumn(2).width = 15;
    worksheet.getColumn(3).width = 15;
    worksheet.getColumn(4).width = 15;
    worksheet.getColumn(5).width = 15;
    worksheet.getColumn(6).width = 15;
    worksheet.getColumn(7).width = 15;
    worksheet.getColumn(8).width = 15;
    worksheet.getColumn(9).width = 15;
    worksheet.getColumn(10).width = 15;
    worksheet.getColumn(11).width = 15;
    worksheet.getColumn(12).width = 15;
    worksheet.getColumn(13).width = 15;
    worksheet.getColumn(14).width = 15;
    worksheet.getColumn(15).width = 15;
    worksheet.getColumn(16).width = 15;
    worksheet.getColumn(17).width = 15;
    worksheet.getColumn(18).width = 15;

    worksheet.addRow([]);
    workbook.xlsx.writeBuffer().then((data: any) => {
      const blob = new Blob([data], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      FileSaver.saveAs(blob, 'Enterprise Tracking' + EXCEL_EXTENSION);
    });
  }

  index = '';
  commentobj: any = {};

  close(data: any) {
    // console.log(data);
    this.index = '';
  }

  selBlock: any;
  commentopen(item: any, i: any, slblock: any = '') {
    this.index = '';
    console.log('Selected Obj :', item);
    //return
    this.selBlock = slblock + i.toString();
    this.index = i.toString() + slblock;

    var lblName = 'Salesperson';
   

    this.commentobj = {
      TYPE: lblName,
      NAME: lblName,
      STORES: 'SALES PERSON',
      STORENAME: lblName,
      Month: '',
      ModuleId: '94',
      ModuleRef: 'SPC',
      state: 1,
      indexval: i,
      mainCat: item.SP1,
    };

  }

  addcmt(data: any) {
    // console.log('Checking Add cmt  : ', data);
    if (data == 'A') {
      this.index = '';
      const DetailsSF = this.ngbmodal.open({
        size: 'xl',
        backdrop: 'static',
      });
      // myObject['skillItem2'] = 15;
      this.commentobj['state'] = 0;
      (DetailsSF.componentInstance.SFComments = this.commentobj),
        DetailsSF.result.then(
          (data) => {
            // console.log(data);
          },
          (reason) => {
            // console.log(reason);

            if (reason == 'O') {
              this.commentobj['state'] = 1;
              this.index = this.commentobj['indexval'];
            } else {
              this.commentobj['state'] = 1;
              this.index = this.commentobj['indexval'];
            }
            // // on dismiss

            // const Data = {
            //   state: true,
            // };
            // this.apiSrvc.setBackgroundstate({ obj: Data });
            this.GetEnterpriseTracking('');
          }
        );
    }
  }
  generatePDF() {
    this.spinner.show();
    const printContents =
      document.getElementById('EnterpriseTracking')!.innerHTML;
    const iframe = document.createElement('iframe');

    // Make the iframe invisible
    iframe.style.position = 'absolute';
    iframe.style.width = '0px';
    iframe.style.height = '0px';
    iframe.style.border = 'none';

    document.body.appendChild(iframe);

    const doc = iframe.contentWindow?.document;
    if (!doc) {
      console.error('Failed to create iframe document');
      return;
    }

    doc.open();
    doc.write(`
    <html>
    <head>
    <title>Enterprise Tracking</title>
                 <style>
                 @font-face {
                  font-family: 'GothamBookRegular';
                  src: url('assets/fonts/Gotham\ Book\ Regular.otf') format('otf'), /* Chrome 6+, Firefox 3.6+, IE 9+, Safari 5.1+ */
                       url('assets/fonts/Gotham\ Book\ Regular.otf') format('opentype'); /* Chrome 4+, Firefox 3.5, Opera 10+, Safari 3—5 */
                }
                @font-face {
                  font-family: 'Roboto';
                  src: url('assets/fonts/Roboto-Regular.ttf') format('ttf'), /* Chrome 6+, Firefox 3.6+, IE 9+, Safari 5.1+ */
                       url('assets/fonts/Roboto-Regular.ttf') format('truetype'); /* Chrome 4+, Firefox 3.5, Opera 10+, Safari 3—5 */
                }
                @font-face {
                  font-family: 'RobotoBold';
                  src: url('assets/fonts/Roboto-Bold.ttf') format('ttf'), /* Chrome 6+, Firefox 3.6+, IE 9+, Safari 5.1+ */
                       url('assets/fonts/Roboto-Bold.ttf') format('truetype'); /* Chrome 4+, Firefox 3.5, Opera 10+, Safari 3—5 */
                }
                .justify-content-between {
                  justify-content: space-between !important;
              }
              .d-flex {
                  display: flex !important;
              }
                .negative {
                  color: red;
                }
                .bg-white{
                  background: #ffffff !important;
              }
                .performance-scorecard .table>:not(:first-child){
                  border-top:0px solid #ffa51a
                }
                .performance-scorecard .table{
                  text-align: center;
                  text-transform: capitalize;
                  border: transparent;
            
                  width: 100%;
                }
                .performance-scorecard .table th,
                .performance-scorecard .table td{
                  white-space: nowrap;
                  vertical-align: top;
                }
                .performance-scorecard .table  th:first-child,
                .performance-scorecard .table td:first-child{
                  position: sticky;
                  left: 0;
                  z-index: 1;
                  background-color: #337ab7;
                }
                .performance-scorecard .table tr:nth-child(odd) td:first-child,
                .performance-scorecard .table tr:nth-child(odd) td:nth-child(2){
                  background-color: #e9ecef;
                 } 
                 .performance-scorecard .table tr:nth-child(even) td:first-child,
                 .performance-scorecard .table tr:nth-child(even) td:nth-child(2){
                  background-color: #ffffff;
                 }
                 .performance-scorecard .table tr:nth-child(odd){
                  background-color: #e9ecef;
                 }
                 .performance-scorecard .table tr:nth-child(even){
                  background-color: #ffffff;
                 }
                 .performance-scorecard .table .spacer {
                  // width: 50px !important;
                  background-color: #cfd6de  !important;
                  border-left: 1px solid #cfd6de !important;
                  border-bottom: 1px solid #cfd6de !important;
                  border-top: 1px solid #cfd6de !important;
                 }
                 .performance-scorecard .table .bdr-rt{
                  border-right: 1px solid #abd0ec;
                 }
                 .performance-scorecard .table thead{
                  position: sticky;
                  top: 0;
                  z-index: 99;
                  font-family: 'FaktPro-Bold';
                  font-size: 0.8rem;
                 }
                 .performance-scorecard .table thead th{
                  padding: 5px 10px;
                  margin: 0px;
                  width: 200px;
                  white-space: normal; 
                 }
                 .performance-scorecard .table thead tr:nth-child(1) { 
                  background-color: #fff !important;
                  color: #000;
                  text-transform: uppercase;
                  border-bottom: #cfd6de;
                 }
                 .performance-scorecard .table thead  tr:nth-child(2) { 
                  background-color: #337ab7 !important;
                  color: #fff;
                  text-transform: uppercase;
                  border-bottom: #cfd6de;
                  box-shadow: inset 0 1px 0 0 #cfd6de;
                 } 
                 .performance-scorecard .table thead tr:nth-child(3) {
                  background-color: #337ab7 !important;
                  color: #fff;
                  text-transform: uppercase; 
                  border-bottom: #cfd6de;
                  box-shadow: inset 0 1px 0 0 #cfd6de;
                }
                .performance-scorecard .table tbody{
                  font-family: 'FaktPro-Normal';
                  font-size: .9rem;
                 }  
                 .performance-scorecard .table tbody td:not(:first-child){
                  text-align: center;
                 } 
                 .performance-scorecard .table tbody td{
                  padding:2px 10px;
                  margin: 0px; 
                  border: 1px solid #cfd6de 
                 } 
                 .performance-scorecard .table tbody tr{
                  border-bottom: 1px solid #cfd6de;
                  border-left: 1px solid #cfd6de
                 }
                 .performance-scorecard .table tbody td:first-child{
                  text-align: start;
                  box-shadow:inset -1px 0 0 0 #cfd6de ;
                 }
                 .performance-scorecard .table tbody .text-bold{
                  font-family: 'FaktPro-Bold';
                 }
                 .performance-scorecard .table tbody .darkred-bg{ 
                  background-color: #282828 !important;
                  color: #fff; 
                 }
                 .performance-scorecard .table tbody .lightblue-bg{
                  background-color: #94b6d1 !important;
                  }
                  .performance-scorecard .table tbody  .gold-bg{ 
                   background-color: #ffa51a;
                   color: #fff;
                   }
                   .performance-scorecard .table tbody .extra-padding{
                    padding-left: 1.5rem;
                  }
              
                
              
                 </style>
  </head>
  <body id='content'>
        ${printContents}
        </body>
          </html>`);
    doc.close();

    const div = doc.getElementById('content');
    const options = {
      logging: true,
      allowTaint: false,
      useCORS: true,
    };
    if (!div) {
      console.error('Element not found');
      return;
    }
    html2canvas(div, options)
      .then((canvas) => {
        let imgWidth = 286;
        let pageHeight = 204;
        let imgHeight = (canvas.height * imgWidth) / canvas.width;
        let heightLeft = imgHeight;
        const contentDataURL = canvas.toDataURL('image/png');
        let pdfData = new jsPDF('l', 'mm', 'a4', true);
        let position = 5;

        function addExtraDataToPage(pdf: any, extraData: any, positionY: any) {
          pdf.text(extraData, 10, positionY);
        }

        function addPageAndImage(pdf: any, contentDataURL: any, position: any) {
          pdf.addPage();
          pdf.addImage(
            contentDataURL,
            'PNG',
            5,
            position,
            imgWidth,
            imgHeight,
            undefined,
            'FAST'
          );
        }

        pdfData.addImage(
          contentDataURL,
          'PNG',
          5,
          position,
          imgWidth,
          imgHeight,
          undefined,
          'FAST'
        );
        addExtraDataToPage(pdfData, '', position + imgHeight + 10);
        heightLeft -= pageHeight;

        while (heightLeft >= 0) {
          position = heightLeft - imgHeight;
          addPageAndImage(pdfData, contentDataURL, position);
          addExtraDataToPage(pdfData, '', position + imgHeight + 10);
          heightLeft -= pageHeight;
        }

        return pdfData;
      })
      .then((doc) => {
        doc.save('Enterprise Tracking.pdf');
        // popupWin.close();
        this.spinner.hide();
      });
  }
  sendEmailData(Email: any, notes: any, from: any) {
    this.spinner.show();
    const printContents =
      document.getElementById('EnterpriseTracking')!.innerHTML;
    const iframe = document.createElement('iframe');

    // Make the iframe invisible
    iframe.style.position = 'absolute';
    iframe.style.width = '0px';
    iframe.style.height = '0px';
    iframe.style.border = 'none';

    document.body.appendChild(iframe);

    const doc = iframe.contentWindow?.document;
    if (!doc) {
      console.error('Failed to create iframe document');
      return;
    }

    doc.open();
    doc.write(`
        <html>
            <head>
            <title>Enterprise Tracking</title>
                 <style>
                 @font-face {
                  font-family: 'GothamBookRegular';
                  src: url('assets/fonts/Gotham\ Book\ Regular.otf') format('otf'), /* Chrome 6+, Firefox 3.6+, IE 9+, Safari 5.1+ */
                       url('assets/fonts/Gotham\ Book\ Regular.otf') format('opentype'); /* Chrome 4+, Firefox 3.5, Opera 10+, Safari 3—5 */
                }
                @font-face {
                  font-family: 'Roboto';
                  src: url('assets/fonts/Roboto-Regular.ttf') format('ttf'), /* Chrome 6+, Firefox 3.6+, IE 9+, Safari 5.1+ */
                       url('assets/fonts/Roboto-Regular.ttf') format('truetype'); /* Chrome 4+, Firefox 3.5, Opera 10+, Safari 3—5 */
                }
                @font-face {
                  font-family: 'RobotoBold';
                  src: url('assets/fonts/Roboto-Bold.ttf') format('ttf'), /* Chrome 6+, Firefox 3.6+, IE 9+, Safari 5.1+ */
                       url('assets/fonts/Roboto-Bold.ttf') format('truetype'); /* Chrome 4+, Firefox 3.5, Opera 10+, Safari 3—5 */
                }
                .justify-content-between {
                  justify-content: space-between !important;
              }
              .d-flex {
                  display: flex !important;
              }
                .negative {
                  color: red;
                }
                .bg-white{
                  background: #ffffff !important;
              }
                .performance-scorecard .table>:not(:first-child){
                  border-top:0px solid #ffa51a
                }
                .performance-scorecard .table{
                  text-align: center;
                  text-transform: capitalize;
                  border: transparent;
            
                  width: 100%;
                }
                .performance-scorecard .table th,
                .performance-scorecard .table td{
                  white-space: nowrap;
                  vertical-align: top;
                }
                .performance-scorecard .table  th:first-child,
                .performance-scorecard .table td:first-child{
                  position: sticky;
                  left: 0;
                  z-index: 1;
                  background-color: #337ab7;
                }
                .performance-scorecard .table tr:nth-child(odd) td:first-child,
                .performance-scorecard .table tr:nth-child(odd) td:nth-child(2){
                  background-color: #e9ecef;
                 } 
                 .performance-scorecard .table tr:nth-child(even) td:first-child,
                 .performance-scorecard .table tr:nth-child(even) td:nth-child(2){
                  background-color: #ffffff;
                 }
                 .performance-scorecard .table tr:nth-child(odd){
                  background-color: #e9ecef;
                 }
                 .performance-scorecard .table tr:nth-child(even){
                  background-color: #ffffff;
                 }
                 .performance-scorecard .table .spacer {
                  // width: 50px !important;
                  background-color: #cfd6de  !important;
                  border-left: 1px solid #cfd6de !important;
                  border-bottom: 1px solid #cfd6de !important;
                  border-top: 1px solid #cfd6de !important;
                 }
                 .performance-scorecard .table .bdr-rt{
                  border-right: 1px solid #abd0ec;
                 }
                 .performance-scorecard .table thead{
                  position: sticky;
                  top: 0;
                  z-index: 99;
                  font-family: 'FaktPro-Bold';
                  font-size: 0.8rem;
                 }
                 .performance-scorecard .table thead th{
                  padding: 5px 10px;
                  margin: 0px;
                  width: 200px;
                  white-space: normal; 
                 }
                 .performance-scorecard .table thead tr:nth-child(1) { 
                  background-color: #fff !important;
                  color: #000;
                  text-transform: uppercase;
                  border-bottom: #cfd6de;
                 }
                 .performance-scorecard .table thead  tr:nth-child(2) { 
                  background-color: #337ab7 !important;
                  color: #fff;
                  text-transform: uppercase;
                  border-bottom: #cfd6de;
                  box-shadow: inset 0 1px 0 0 #cfd6de;
                 } 
                 .performance-scorecard .table thead tr:nth-child(3) {
                  background-color: #337ab7 !important;
                  color: #fff;
                  text-transform: uppercase; 
                  border-bottom: #cfd6de;
                  box-shadow: inset 0 1px 0 0 #cfd6de;
                }
                .performance-scorecard .table tbody{
                  font-family: 'FaktPro-Normal';
                  font-size: .9rem;
                 }  
                 .performance-scorecard .table tbody td:not(:first-child){
                  text-align: center;
                 } 
                 .performance-scorecard .table tbody td{
                  padding:2px 10px;
                  margin: 0px; 
                  border: 1px solid #cfd6de 
                 } 
                 .performance-scorecard .table tbody tr{
                  border-bottom: 1px solid #cfd6de;
                  border-left: 1px solid #cfd6de
                 }
                 .performance-scorecard .table tbody td:first-child{
                  text-align: start;
                  box-shadow:inset -1px 0 0 0 #cfd6de ;
                 }
                 .performance-scorecard .table tbody .text-bold{
                  font-family: 'FaktPro-Bold';
                 }
                 .performance-scorecard .table tbody .darkred-bg{ 
                  background-color: #282828 !important;
                  color: #fff; 
                 }
                 .performance-scorecard .table tbody .lightblue-bg{
                  background-color: #94b6d1 !important;
                  }
                  .performance-scorecard .table tbody  .gold-bg{ 
                   background-color: #ffa51a;
                   color: #fff;
                   }
                   .performance-scorecard .table tbody .extra-padding{
                    padding-left: 1.5rem;
                  }
              
                
              
                 </style>
            </head>
            <body id='content'>
                ${printContents}
            </body>
        </html>
    `);

    doc.close();

    const div = doc.getElementById('content');
    if (!div) {
      console.error('Element not found');
      return;
    }

    const options = {
      logging: true,
      allowTaint: false,
      useCORS: true,
      scale: 1, // Adjust scale to fit the page better
    };

    html2canvas(div, options)
      .then((canvas) => {
        let imgWidth = 286;
        let pageHeight = 204;
        let imgHeight = (canvas.height * imgWidth) / canvas.width;
        let heightLeft = imgHeight;
        const contentDataURL = canvas.toDataURL('image/png');
        let pdfData = new jsPDF('l', 'mm', 'a4', true);
        let position = 5;
        pdfData.addImage(
          contentDataURL,
          'PNG',
          5,
          position,
          imgWidth,
          imgHeight,
          undefined,
          'FAST'
        );
        heightLeft -= pageHeight;
        while (heightLeft >= 0) {
          position = heightLeft - imgHeight;
          pdfData.addPage();
          pdfData.addImage(
            contentDataURL,
            'PNG',
            5,
            position,
            imgWidth,
            imgHeight,
            undefined,
            'FAST'
          );
          heightLeft -= pageHeight;
        }

        const pdfBlob = pdfData.output('blob');
        const pdfFile = this.blobToFile(pdfBlob, 'Enterprise Tracking.pdf');
        const formData = new FormData();
        formData.append('to_email', Email);
        formData.append('subject', 'Enterprise Tracking');
        formData.append('file', pdfFile);
        formData.append('notes', notes);
        formData.append('from', from);
        this.apiSrvc.postmethod(this.comm.routeEndpoint + 'mail', formData).subscribe(
          (res: any) => {
            console.log('Response:', res);
            if (res.status === 200)  {
             
              this.toast.show('Invalid Details.', 'danger', 'Error');
            }
          },
          (error) => {
            console.error('Error:', error);
          }
        );
      })
      .catch((error) => {
        console.error('html2canvas error:', error);
      })
      .finally(() => {
        this.spinner.hide();
        // popupWin.close();
      });
  }
  public blobToFile = (theBlob: Blob, fileName: string): File => {
    return new File([theBlob], fileName, {
      lastModified: new Date().getTime(),
      type: theBlob.type,
    });
  };

  GetPrintData() {
    window.print();
  }

}
