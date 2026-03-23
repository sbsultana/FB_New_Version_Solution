import { Component, ElementRef, Injector, ViewChild, HostListener } from '@angular/core';
import { Api } from '../../../../Core/Providers/Api/api';
import { common } from '../../../../common';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { FormsModule } from '@angular/forms';
import { NgbActiveModal, NgbModal } from '@ng-bootstrap/ng-bootstrap';
import { NgxSpinnerService } from 'ngx-spinner';
import { CurrencyPipe, DatePipe, formatDate } from '@angular/common';
import { Title } from '@angular/platform-browser';
import { Subscription } from 'rxjs';
import { environment } from '../../../../../environments/environment';
import { Workbook, Row } from 'exceljs';
import FileSaver from 'file-saver';
import html2canvas from 'html2canvas';
import jsPDF from 'jspdf';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
const EXCEL_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
const EXCEL_EXTENSION = '.xlsx';
import numeral from 'numeral';

import { BsDatepickerConfig, BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { FinancialsummaryDetails } from '../financialsummary-details/financialsummary-details';
import { Stores } from '../../../../CommonFilters/stores/stores';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [SharedModule, BsDatepickerModule, Stores],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss'
})
export class Dashboard {
  Report: any = '';
  dynamicTitle = 'Financial Summary';
  FromDate: any;
  ToDate: any;

  reportOpenSub!: Subscription;
  reportGetting!: Subscription;
  Pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;

  Current_Date: any;
  LMY_Date: any;
  LM_Date: any;

  FSData: any = [];
  Month: any = '';
  showAdditionalRows: boolean = false;
  hiddenRows: boolean[] = [];

  EBITDAdata: any = [];
  ETBudgetData: any = [];
  ETDealerData: any = [];

  SelectedTab: any = ['FinancialSummary'];
  Filter: any = 'FinancialSummary';
  SubFilter: any = '';
  StoreName: any = 'All Stores';
  StoreValues: any = [];
  groups: any = 1;

  NoData: boolean = false;
  date: any = new Date();
  roleId: any;
  PresentDayDate: string;
  header: any = [
    {
      type: 'Bar',
      StoreValues: this.StoreValues,
      Month: this.Month,
      groups: this.groups,
    },
  ];
  activePopover: number = -1;
  storeIds: any = '0';
  popup: any = [{ type: 'Popup' }];
  pdfStyleService: any;
  selectedDate: Date = new Date();
  currentMonth!: Date;
  groupsArray: any = [];
  storename: any = ''
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  groupId: any = 0;
  stores: any = [];
  storesFilterData: any = {
    'groupsArray': this.groupsArray, 'groupId': this.groupId, 'storesArray': this.stores, 'storeids': '1', 'type': 'M', 'others': 'N',
    'groupName': this.groupName, 'storename': this.storename, storecount: null, 'storedisplayname': this.storedisplayname
  };

  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent) {
    const clickedInside = (event.target as HTMLElement).closest('.dropdown-toggle, .reportstores-card , .timeframe');
    if (!clickedInside) {
      this.activePopover = -1;
    }
  }
  constructor(
    public apiSrvc: Api,
    private ngbmodal: NgbModal,
    private ngbmodalActive: NgbActiveModal,
    private spinner: NgxSpinnerService,
    private datepipe: DatePipe,
    private title: Title,
    private comm: common,
    private toast: ToastService,
    private injector: Injector,
    public shared: Sharedservice,
  ) {
    if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
      // this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.ustores.split(',')
      this.groupId = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Preferences
      this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids.split(',')
    }
    console.log('store displayname', this.storedisplayname);
    if (this.shared.common.groupsandstores.length > 0) {
      this.groupsArray = this.shared.common.groupsandstores.filter((val: any) => val.sg_id != this.shared.common.reconID);
      this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
      this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_Name : this.groupName = ''
      this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
      this.getStoresandGroupsValues()
    }
    this.title.setTitle(this.comm.titleName + '-Financial Summary');
    const data = { title: this.dynamicTitle, stores: '2' };
    const lastMonth = new Date();
    let today = new Date();
    if (today.getDate() < 5) {
      this.date = new Date(lastMonth.setMonth(lastMonth.getMonth() - 1));
    } else {
      this.date = new Date(lastMonth.setMonth(lastMonth.getMonth()));
    }

    if (localStorage.getItem('Fav') != 'Y') {
      const data = {
        title: 'Financial Summary',
        path1: '',
        path2: '',
        path3: '',
        Month: this.date,
        stores: this.StoreValues.toString(),
        store: this.storeIds,
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
      this.Month =
        this.date.toString().substr(8, 2) +
        '-' +
        this.date.toString().substr(4, 3) +
        '-' +
        this.date.toString().substr(11, 4);
      this.selectedDate = this.date;
      this.currentMonth = this.selectedDate;
      this.GetData(this.currentMonth);
    }

    const format = 'ddMMyyyy';
    const locale = 'en-US';
    const myDate = new Date();
    const formattedDate = formatDate(myDate, format, locale);
    this.PresentDayDate = formattedDate;
  }

  ngOnInit(): void {
    this.roleId = localStorage.getItem('roleId');
    // var curl =
    //   'https://fbxtract.axelautomotive.com/favouritereports/GetFinancialSummaryReportfb';
    // this.apiSrvc.logSaving(curl, {}, '', 'Success', 'Financial Summary');
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
      this.GetData(this.currentMonth);
    }
  }

  FsData: any = [];

  Scrollpercent: any = 0;
  updateVerticalScroll(event: any): void {
    const scrollDemo = document.querySelector('#scrollcent') as HTMLElement;
    this.Scrollpercent = Math.round(
      (event.target.scrollTop /
        (event.target.scrollHeight - scrollDemo.clientHeight)) *
      100
    );
  }


  GetData(date: Date) {
    this.Month =
      date.toString().substr(8, 2) +
      '-' +
      date.toString().substr(4, 3) +
      '-' +
      date.toString().substr(11, 4);

    this.Current_Date =
      date.toString().substr(8, 2) +
      '-' +
      date.toString().substr(4, 3) +
      '-' +
      date.toString().substr(11, 4);
    this.LMY_Date =
      date.toString().substr(8, 2) +
      '-' +
      date.toString().substr(4, 3) +
      '-' +
      (date.getFullYear() - 1);

    // console.log(this.LMY_Date);

    let LM_StartDate = new Date();
    let LMDate =
      date.toString().substr(8, 2) +
      '-' +
      date.toString().substr(4, 3) +
      '-' +
      date.toString().substr(11, 4);

    let sample = new Date(LMDate.replace(/-/g, '/'));
    LM_StartDate = new Date(sample.setMonth(sample.getMonth() - 1));

    this.LM_Date =
      date.toString().substr(8, 2) +
      '-' +
      LM_StartDate.toString().substr(4, 3) +
      '-' +
      LM_StartDate.toString().substr(11, 4);

    this.FSData = [];
    this.spinner.show();
    const DateToday = this.datepipe.transform(
      new Date(date),
      'yyyy-MM-dd'
    );
    let Obj = {
      as_Id: this.storeIds.toString(),
      SalesDate: DateToday,
      // UserID: 0,
    };
    console.log(Obj);
    const curl = environment.apiUrl + this.comm.routeEndpoint + 'GetFinancialSummaryReportfb';
    this.apiSrvc
      .postmethod(this.comm.routeEndpoint + 'GetFinancialSummaryReportfb', Obj)
      .subscribe(
        (res) => {
          const currentTitle = document.title;
          this.apiSrvc.logSaving(curl, {}, '', res.message, currentTitle);
          if (res.status == 200) {
            this.FSData = res.response;
            const newId = Math.floor(Math.random() * 1000) + 1;
            const newObj = {
              id: newId,
              LABLE1: 'Variable Operations',
              MainTitle: 'MainHeader',
            };
            this.FSData.unshift(newObj);
            console.log('Array Values', this.FSData);
            const Index = this.FSData.findIndex(
              (obj: any) => obj.LABLE1 === 'Variable Gross'
            );
            if (Index !== -1 && Index < this.FSData.length - 1) {
              const newId = Math.floor(Math.random() * 1000) + 1;
              const newObj = {
                id: newId,
                LABLE1: 'Fixed Operations',
                MainTitle: 'MainHeader',
              };
              this.FSData.splice(Index + 1, 0, newObj);
            }
            const IndexOne = this.FSData.findIndex(
              (obj: any) => obj.LABLE1 === 'Total Store Gross'
            );
            if (IndexOne !== -1 && IndexOne < this.FSData.length - 1) {
              const newId = Math.floor(Math.random() * 1000) + 1;
              const newObj = {
                id: newId,
                LABLE1: 'Expenses',
                MainTitle: 'MainHeader',
              };
              this.FSData.splice(IndexOne + 1, 0, newObj);
            }
            const IndexTwo = this.FSData.findIndex(
              (obj: any) => obj.LABLE1 === 'Operating Profit'
            );
            if (IndexTwo !== -1 && IndexTwo < this.FSData.length - 1) {
              const newId = Math.floor(Math.random() * 1000) + 1;
              const newObj = {
                id: newId,
                LABLE1: 'Net Adjustments/Other Income',
                MainTitle: 'MainHeader',
              };
              this.FSData.splice(IndexTwo + 1, 0, newObj);
            }
            const IndexThree = this.FSData.findIndex(
              (obj: any) => obj.LABLE1 === 'Net to Gross'
            );
            if (IndexThree !== -1 && IndexThree < this.FSData.length - 1) {
              const newId = Math.floor(Math.random() * 1000) + 1;
              const newObj = {
                id: newId,
                LABLE1: 'Memo: Super Gross',
                MainTitle: 'MainHeader',
              };
              this.FSData.splice(IndexThree + 1, 0, newObj);
            }
            this.FSData.forEach((val: any) => {
              if (
                val.LABLE1 == 'Unit Retail Sales' ||
                val.LABLE1 == 'Pure Gross' ||
                val.LABLE1 == 'Variable Gross' ||
                val.LABLE1 == 'Total Store Gross' ||
                val.LABLE1 == 'Selling Gross' ||
                val.LABLE1 == 'Selling Gross%' ||
                val.LABLE1 == 'Total Expenses' ||
                val.LABLE1 == 'Operating Profit' ||
                val.LABLE1 == 'Net Adds/Deducts' ||
                val.LABLE1 == 'Net Profit' ||
                val.LABLE1 == 'Total Store Super Gross'
              ) {
                val.Fs_Titles = 'FontBold';
              } else if (
                val.LABLE1 == 'New' ||
                val.LABLE1 == 'Used' ||
                val.LABLE1 == 'Service' ||
                val.LABLE1 == 'Parts'
              ) {
                val.Fs_Titles = 'PaddingLeft';
              }
            });
            this.FSData.forEach((val: any, index: any) => {
              if (val.LABLE1 == 'Selling Gross%') {
                this.showAdditionalRows = !this.showAdditionalRows;
                for (
                  let i = index + 1;
                  i < index + 5 && i < this.FSData.length;
                  i++
                ) {
                  this.hiddenRows[i] = true;
                }
              }
            });
            this.spinner.hide();
            if (this.FSData.length > 0) {
              this.NoData = false;
            } else {
              this.NoData = true;
            }
          } else {

            this.toast.show('Invalid Details.', 'danger', 'Error');
          }
        },
        (error) => {
          console.log(error);
        }
      );
  }

  toggleRows(index: number): void {
    console.log(index);
    if (this.FSData[index].LABLE1 === 'Selling Gross%') {
      this.showAdditionalRows = !this.showAdditionalRows;
      for (let i = index + 1; i < index + 5 && i < this.FSData.length; i++) {
        this.hiddenRows[i] = !this.hiddenRows[i];
        console.log(this.hiddenRows);
      }
    }
  }
  getStyle(value: any, valueType: string) {
    if (value === null || value === '' || value === 0 || value === '0') {
      return { 'justify-content': 'center' };
    }
    switch (valueType) {
      case '$':
        return { 'justify-content': 'space-between' };
      case '#':
        return { 'justify-content': 'flex-end' };
      case '%':
        return { 'justify-content': 'center' };
      default:
        return { 'justify-content': 'initial' };
    }
  }
  formatValue(value: any, valueType: string): string {
    if (value === null || value === undefined || value === 0 || value === '0') {
      return '-';
    } else if (valueType === '#') {
      return typeof value === 'number' ? numeral(value).format('0,0') : value;
    } else if (valueType === '$') {
      return typeof value === 'number' ? numeral(value).format('0,0') : value;
    } else if (valueType === '%') {
      return typeof value === 'number' ? value.toLocaleString() + '%' : value;
    } else {
      return value;
    }
  }

  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    }
    return false;
  }

  // StoreValues: any = 0;
  FsClick: any;
  block: any = '';
  ngAfterViewInit(): void {
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.shared.common.pageName == 'Financial Summary') {
        if (res.obj.storesData != undefined) {
          this.groupsArray = res.obj.storesData;
          this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
          this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_name : this.groupName = ''
          this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
          this.getStoresandGroupsValues()
        }
      }
    })
    this.reportOpenSub = this.apiSrvc.GetReportOpening().subscribe((res) => {
      if (this.reportOpenSub != undefined) {
        if (res.obj.Module == 'Financial Summary') {
          document.getElementById('report')?.click();
        }
      }
    });
    this.reportGetting = this.apiSrvc.GetReports().subscribe((data) => {
      console.log(data);
      if (this.reportGetting != undefined) {
        if (data.obj.Reference == 'Financial Summary') {
          if (data.obj.header == undefined) {
            this.date = data.obj.month;
            this.Month = data.obj.month;
            this.StoreValues = data.obj.storeValues;
            this.StoreName = data.obj.Sname;
            this.groups = data.obj.groups;
          } else {
            if (data.obj.header == 'Yes') {
              this.StoreValues = data.obj.storeValues;
            }
          }
          if (this.StoreValues != '') {
            this.GetData(this.currentMonth);
          } else {
            this.NoData = true;
            this.FSData = [];
          }
          const headerdata = {
            title: 'Financial Summary',
            path1: '',
            path2: '',
            path3: '',
            Month: new Date(this.Month),
            stores: this.StoreValues,
            groups: this.groups,
          };
          this.apiSrvc.SetHeaderData({
            obj: headerdata,
          });
          this.header = [
            {
              type: 'Bar',
              StoreValues: this.StoreValues,
              Month: new Date(this.Month),
              groups: this.groups,
            },
          ];
        }
      }
    });
    this.excel = this.apiSrvc.getExportToExcelAllReports().subscribe((res: { obj: { state: boolean; title: string; }; }) => {
      this.FsClick = res.obj.state;
      if (this.excel != undefined) {
        if (res.obj.title == 'Financial Summary') {
          if (res.obj.state == true) {
            this.exportToExcelFinancialSummary();
          }
        }
      }
    });
    this.print = this.apiSrvc.getExportToPrintAllReports().subscribe((res: { obj: { state: boolean; title: string; }; }) => {
      if (this.print != undefined) {
        if (res.obj.title == 'Financial Summary') {
          if (res.obj.state == true) {
            this.GetPrintData();
          }
        }
      }
    });
    this.Pdf = this.apiSrvc.getExportToPDFAllReports().subscribe((res: { obj: { state: boolean; title: string; }; }) => {
      if (this.Pdf != undefined) {
        if (res.obj.title == 'Financial Summary') {
          if (res.obj.state == true) {
            this.generatePDF();
          }
        }
      }
    });
    this.email = this.apiSrvc.getExportToEmailPDFAllReports().subscribe((res: { obj: { state: boolean; title: string; }; }) => {
      if (this.email != undefined) {
        if (res.obj.title == 'Financial Summary') {
          if (res.obj.state == true) {
            // this.sendEmailData(res.obj.Email, res.obj.notes, res.obj.from);
          }
        }
      }
    });
  }
  ngOnDestroy() {
    // this.reportOpenSub.unsubscribe()
    // this.reportGetting.unsubscribe()
    // this.Pdf.unsubscribe()
    // this.print.unsubscribe()
    // this.email.unsubscribe()
    // this.excel.unsubscribe()
    if (this.reportOpenSub != undefined) {
      this.reportOpenSub.unsubscribe()
    }
    if (this.reportGetting != undefined) {
      this.reportGetting.unsubscribe()
    }
    if (this.excel != undefined) {
      this.excel.unsubscribe()
    }
    if (this.Pdf != undefined) {
      this.Pdf.unsubscribe()
    }
    if (this.print != undefined) {
      this.print.unsubscribe()
    }
    if (this.email != undefined) {
      this.email.unsubscribe()
    }
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

  reportOpen(temp: any) {
    this.ngbmodalActive = this.ngbmodal.open(temp, {
      size: 'xl',
      backdrop: 'static',
    });
  }
  SubSelectedTab1: any = [];

  openMenu(Object: any, LatestDate: any) {
    const DetailsFs = this.ngbmodal.open(FinancialsummaryDetails, {
      size: 'xl',
      backdrop: 'static',
      injector: Injector.create({
        providers: [
          { provide: CurrencyPipe, useClass: CurrencyPipe }
        ],
        parent: this.injector
      })
    });

    DetailsFs.componentInstance.Fsdetails = {
      TYPE: Object.LABLE1Val,
      NAME: Object.LABLE1,
      STORES: this.storeIds,
      LatestDate: LatestDate,
      Group: this.groups,
    };
  }
  GraphHeadings: any;
  GraphData: any;
  ValueFormat: any;
  openGraph(Object: any, Current_Date: any, LMY_Date: any, LM_Date: any) {
    // console.log(Object, Current_Date, LMY_Date, LM_Date);
    const CurrentMonth = this.datepipe.transform(Current_Date, 'MMMM yyyy');
    const LastYearMonth = this.datepipe.transform(LMY_Date, 'MMMM yyyy');
    const LastMonth = this.datepipe.transform(LM_Date, 'MMMM yyyy');
    const CurrentYear = this.datepipe.transform(Current_Date, 'yyyy');
    const LastYear = this.datepipe.transform(LMY_Date, 'yyyy');
    this.GraphHeadings = [
      CurrentMonth,
      LastYearMonth,
      LastMonth,
      CurrentYear + ' Average',
      LastYear + ' Average',
    ];
    this.GraphData = [
      Object.MTD,
      Object.LMY_MTD,
      Object.LM_MTD,
      Object.YTD,
      Object.LY_YTD,
    ];
    if (
      Object.LABLE1 == 'New Units' ||
      Object.LABLE1 == 'Pre-Owned Units' ||
      Object.LABLE1 == 'Wholesale Units' ||
      Object.LABLE1 == 'Unit Retail Sales'
    ) {
      this.ValueFormat = 'Number';
    } else if (
      Object.LABLE1 == 'Selling Gross%' ||
      Object.LABLE1 == 'New' ||
      Object.LABLE1 == 'Used' ||
      Object.LABLE1 == 'Service' ||
      Object.LABLE1 == 'Parts' ||
      Object.LABLE1 == 'Net to Sales' ||
      Object.LABLE1 == 'Net to Gross'
    ) {
      this.ValueFormat = 'Percentage';
    } else {
      this.ValueFormat = 'Currancy';
    }
    const DetailsSF = this.ngbmodal.open({
      size: 'xl',
      backdrop: 'static',
    });
    DetailsSF.componentInstance.FSgraphdetails = {
      Header: this.GraphHeadings,
      Data: this.GraphData,
      NAME: Object.LABLE1,
      ValueFormat: this.ValueFormat,
      STORES: this.StoreValues,
      LatestDate: this.Month,
      Group: this.groups,
    };
  }
  openfs() {
    const DetailsFs = this.ngbmodal.open({
      size: 'xl',
      backdrop: 'static',
    });
  }

  togglePopover(popoverIndex: number) {
    this.activePopover = this.activePopover === popoverIndex ? -1 : popoverIndex;
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
      title: this.dynamicTitle,
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
  exportToExcelFinancialSummary() {

    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet('Financial Summary');

    /* ================= 1. FILTERS AT TOP ================= */
    const filterRowCount = this.addExcelFiltersSection(worksheet);

    const startRow = filterRowCount + 1;

    /* ================= 2. HEADER ROW 1 ================= */

    worksheet.getCell(`A${startRow}`).value = 'For the period ending';

    worksheet.getCell(`B${startRow}`).value = 'MTD';
    worksheet.mergeCells(`B${startRow}:E${startRow}`);

    worksheet.getCell(`F${startRow}`).value = 'LY MTD';
    worksheet.mergeCells(`F${startRow}:G${startRow}`);

    worksheet.getCell(`H${startRow}`).value = 'LM';
    worksheet.mergeCells(`H${startRow}:I${startRow}`);

    worksheet.getCell(`J${startRow}`).value = 'YTD';
    worksheet.mergeCells(`J${startRow}:L${startRow}`);

    worksheet.getCell(`M${startRow}`).value = 'LY YTD';
    worksheet.mergeCells(`M${startRow}:N${startRow}`);

    worksheet.getRow(startRow).eachCell(cell => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '0554EF' } };
      cell.font = { bold: true, color: { argb: 'FFFFFF' }, name: 'Calibri' };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.border = {
        top: { style: 'thin' },
        bottom: { style: 'thin' },
        left: { style: 'thin' },
        right: { style: 'thin' }
      };
    });

    /* ================= 3. HEADER ROW 2 ================= */

    const header2 = [
      this.datepipe.transform(this.currentMonth, 'MMMM yyyy'),

      this.datepipe.transform(this.currentMonth, 'MMMM'),
      'Pace',
      'Budget',
      'Variance',

      this.datepipe.transform(this.LMY_Date, 'MMMM'),
      'Variance',

      this.datepipe.transform(this.LM_Date, 'MMMM'),
      'Variance',

      this.datepipe.transform(this.currentMonth, 'yyyy'),
      'Budget',
      'Variance',

      this.datepipe.transform(this.LMY_Date, 'yyyy'),
      'Variance'
    ];

    const headerRow2 = worksheet.addRow(header2);

    headerRow2.eachCell(cell => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '4584FF' } };
      cell.font = { bold: true, color: { argb: 'FFFFFF' }, name: 'Calibri' };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.border = {
        top: { style: 'thin' },
        bottom: { style: 'thin' },
        left: { style: 'thin' },
        right: { style: 'thin' }
      };
    });

    /* ================= VARIANCE COLUMN INDEX ================= */

    const varianceColumns: number[] = [];
    header2.forEach((h: any, i) => {
      if (h && h.toString().toLowerCase().includes('variance')) {
        varianceColumns.push(i + 1);
      }
    });

    /* ================= FORMAT FUNCTION ================= */

    const formatRow = (row: Row, data: any, index: number) => {

      const isHeaderRow = data.MainTitle === 'MainHeader';
      const isBold = data.Fs_Titles === 'FontBold';
      const isPadding = data.Fs_Titles === 'PaddingLeft';
      const isSpecial = this.isSpecialRow(data.LABLE1);
      const valueType = data.ValueType;

      const totalCols = 14;

      let bgColor = '';
      if (isHeaderRow) bgColor = 'D9E7FF';
      else bgColor = index % 2 === 0 ? 'F9FBFF' : 'FFFFFF';

      for (let i = 1; i <= totalCols; i++) {
        row.getCell(i);
      }

      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {

        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: bgColor }
        };

        /* ===== FIRST COLUMN ===== */
        if (colNumber === 1) {
          cell.font = {
            name: 'Calibri',
            bold: isBold || isSpecial || isHeaderRow
          };

          cell.alignment = {
            horizontal: 'left',
            vertical: 'middle',
            indent: isHeaderRow ? 1 : (isPadding ? 3 : 2)
          };
          return;
        }

        if (isHeaderRow) {
          cell.value = '';
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
          cell.font = { name: 'Calibri', bold: true };
          return;
        }

        if (cell.value === null || cell.value === undefined || cell.value === '' || cell.value === 0) {
          cell.value = '-';
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
          return;
        }

        const num = Number(cell.value);
        let fontStyle: any = { name: 'Calibri', bold: isBold };

        if (!isNaN(num)) {

          if (valueType === '$') {
            cell.numFmt = '"$" * #,##0; "$" * -#,##0';
          } else if (valueType === '#') {
            cell.numFmt = '#,##0';
          } else if (valueType === '%') {
            cell.numFmt = '0%';
            cell.value = num / 100;
          }

          if (isSpecial && varianceColumns.includes(colNumber) && num < 0) {
            fontStyle.color = { argb: 'FF0000' };
          }
        }

        cell.font = fontStyle;

        cell.alignment = valueType === '%'
          ? { horizontal: 'center', vertical: 'middle' }
          : { horizontal: 'right', vertical: 'middle' };
      });
    };

    /* ================= DATA ================= */

    this.FSData.forEach((data: any, i: number) => {

      const row = worksheet.addRow([
        data.LABLE1,
        data.MTD,
        data.PACE,
        data.BUDGET,
        data.VARIANCE,
        data.LMY_MTD,
        data.LMY_VARIANCE,
        data.LM_MTD,
        data.LM_VARIANCE,
        data.YTD,
        data.CY_BUDGET,
        data.CY_VARIANCE,
        data.LY_YTD,
        data.LY_VARIANCE
      ]);

      formatRow(row, data, i);
    });

    /* ================= BORDERS ================= */

    worksheet.eachRow(row => {
      row.eachCell(cell => {
        cell.border = {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' }
        };
      });
    });

    /* ================= WIDTH ================= */

    worksheet.columns = [
      { width: 35 },
      ...Array(13).fill({ width: 18 })
    ];

    /* ================= FREEZE ================= */

    worksheet.views = [{
      state: 'frozen',
      xSplit: 1,
      ySplit: startRow + 1
    }];

    /* ================= DOWNLOAD ================= */

    workbook.xlsx.writeBuffer().then(data => {
      FileSaver.saveAs(new Blob([data]), 'FinancialSummary.xlsx');
    });
  }
  //Special rows
  isSpecialRow(name: string): boolean {
    return [
      'Total Selling Gross (New)',
      'Total Selling Gross (Used)',
      'Net Income Before Taxes',
      'Total Operating Department Profit',
      'Operating Profit',
      'Net Profit'
    ].includes(name);
  }



  GetPrintData() {
    window.print();
  }

  generatePDF() {
    this.spinner.show();
    const printContents =
      document.getElementById('FinancialSummary')!.innerHTML;
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
    <title>Financial Summary</title>
                 <style>
                 ${this.pdfStyleService.getFinancialSummaryStyles()}
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
        doc.save('Financial Summary.pdf');
        // popupWin.close();
        this.spinner.hide();
      });
  }

  sendEmailData(Email: any, notes: any, from: any) {
    this.spinner.show();
    const printContents =
      document.getElementById('FinancialSummary')!.innerHTML;
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
            <title>Financial Summary</title>
                
                  <style>
                 ${this.pdfStyleService.getFinancialSummaryStyles()}
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
        const pdfFile = this.blobToFile(pdfBlob, 'Financial Summary.pdf');
        const formData = new FormData();
        formData.append('to_email', Email);
        formData.append('subject', 'Financial Summary');
        formData.append('file', pdfFile);
        formData.append('notes', notes);
        formData.append('from', from);
        this.apiSrvc.postmethod(this.comm.routeEndpoint + 'mail', formData).subscribe(
          (res: any) => {
            console.log('Response:', res);
            if (res.status === 200) {

              this.toast.success(res.response);
            } else {

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
  @ViewChild('printSection') printSection!: ElementRef;

  // printDiv() {
  //   const printContents = this.printSection.nativeElement.innerHTML;
  //   const popupWin = window.open('', '_blank', 'width=800,height=600');

  //   // Collect all stylesheets from the current document
  //   const styles = Array.from(document.styleSheets)
  //     .map((styleSheet: any) => {
  //       try {
  //         return Array.from(styleSheet.cssRules)
  //           .map((rule: any) => rule.cssText)
  //           .join('');
  //       } catch (e) {
  //         return ''; // Ignore CORS-restricted stylesheets
  //       }
  //     })
  //     .join('');

  //   popupWin!.document.open();
  //   popupWin!.document.write(`
  //     <html>
  //       <head>
  //         <title>Print Report</title>
  //         <style>${styles}</style>
  //       </head>
  //       <body onload="window.print();window.close()">
  //         ${printContents}
  //       </body>
  //     </html>
  //   `);
  //   popupWin!.document.close();
  // }

  printDiv() {
    const printContents = this.printSection.nativeElement.innerHTML;
    const iframe = document.createElement('iframe');

    iframe.style.position = 'fixed';
    iframe.style.right = '0';
    iframe.style.bottom = '0';
    iframe.style.width = '0';
    iframe.style.height = '0';
    iframe.style.border = 'none';

    document.body.appendChild(iframe);

    // Collect styles
    const styles = Array.from(document.styleSheets)
      .map((styleSheet: any) => {
        try {
          return Array.from(styleSheet.cssRules)
            .map((rule: any) => rule.cssText)
            .join('');
        } catch (e) {
          return '';
        }
      })
      .join('');

    const doc = iframe.contentWindow?.document;
    if (!doc) return;

    doc.open();
    doc.write(`
<html>
<head>
<title>Print Report</title>
<style>
          ${styles}
          @media print {
            body { -webkit-print-color-adjust: exact !important; }
          }
</style>
</head>
<body>
        ${printContents}
</body>
</html>
  `);
    doc.close();

    setTimeout(() => {
      iframe.contentWindow!.focus();
      iframe.contentWindow!.print();
      document.body.removeChild(iframe);
    }, 300);
  }
}


