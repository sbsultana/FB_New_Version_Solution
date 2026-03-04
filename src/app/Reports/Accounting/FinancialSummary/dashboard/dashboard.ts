import { Component, ElementRef, Injector, ViewChild ,HostListener} from '@angular/core';
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
import { Workbook } from 'exceljs';
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
  imports: [SharedModule, BsDatepickerModule,Stores],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss'
})
export class Dashboard {
  Report: any = '';
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
            this.exportAsXLSX();
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
  ExcelStoreNames: any = [];
  exportAsXLSX(): void {
    const CurrentDate = this.datepipe.transform(
      this.Current_Date,
      'MMMM yyyy'
    );
    const CurrentYear = this.datepipe.transform(this.Current_Date, 'yyyy');
    const LMYDate = this.datepipe.transform(this.LMY_Date, 'MMMM yyyy');
    const LastYear = this.datepipe.transform(this.LMY_Date, 'yyyy');

    const LMDate = this.datepipe.transform(this.LM_Date, 'MMMM yyyy');

    const FSData = this.FSData.map((_arrayElement: any) =>
      Object.assign({}, _arrayElement)
    );

    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet('Financial Summary');
    worksheet.views = [
      {
        state: 'frozen',
        ySplit: 12, // Number of rows to freeze (2 means the first two rows are frozen)
        topLeftCell: 'A13', // Specify the cell to start freezing from (in this case, the third row)
        showGridLines: false,
      },
    ];
    worksheet.addRow('');
    const titleRow = worksheet.addRow(['Financial Summary']);
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
      'MM.dd.yyyy h:mm:ss a'
    );
    const DATE_EXTENSION = this.datepipe.transform(new Date(), 'MMddyyyy');
    const Month = this.datepipe.transform(this.Month, 'MMMM yyyy');
    worksheet.addRow([DateToday]).font = {
      name: 'Arial',
      family: 4,
      size: 9,
    };

    const ReportFilter = worksheet.addRow(['Report Filters :']);
    ReportFilter.font = { name: 'Arial', family: 4, size: 10, bold: true };

    const DealType = worksheet.addRow(['Month :']);
    const dealtype = worksheet.getCell('B6');
    dealtype.value = Month;
    dealtype.font = { name: 'Arial', family: 4, size: 9 };
    dealtype.alignment = { vertical: 'middle', horizontal: 'left' };
    DealType.getCell(1).font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
    };
    // const Groups = worksheet.getCell('A7');
    // Groups.value = 'Group :';
    // const groups = worksheet.getCell('B7');
    // groups.value =
    //   'WESTERN AUTO';
    // groups.font = { name: 'Arial', family: 4, size: 9 };
    // groups.alignment = {
    //   vertical: 'middle',
    //   horizontal: 'left',
    //   wrapText: true,
    // };
    // ===== STORE VALUE FOR EXCEL (FROM UI SELECTION) =====

    let storeValue = '';

    const selectedStoreIds: string[] =
      this.storeIds && this.storeIds.length
        ? this.storeIds.map((id: any) => id.toString())
        : [];

    const allStores: any[] = Array.isArray(this.stores) ? this.stores : [];

    // ✅ Bind ONLY selected store names
    storeValue = allStores
      .filter((s: any) => selectedStoreIds.includes(s.ID.toString()))
      .map((s: any) => s.storename.trim())
      .filter(Boolean)
      .join(', ');

    // ✅ Final fallback (safety)
    if (!storeValue && selectedStoreIds.length) {
      storeValue = selectedStoreIds.join(', ');
    }


    worksheet.mergeCells('B8', 'K10');
    const Stores = worksheet.getCell('A8');
    Stores.value = 'Store :';
    const stores = worksheet.getCell('B8');
    stores.value =
      storeValue;
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
    let dateYear = worksheet.getCell('A11');
    dateYear.value = 'For the period ending'
    dateYear.alignment = { vertical: 'middle', horizontal: 'center' };
    dateYear.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
      color: { argb: 'FFFFFF' },
    };
    dateYear.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '0554ef' },
      bgColor: { argb: 'FF0000FF' },
    };
    dateYear.border = { right: { style: 'thin' } };

    worksheet.mergeCells('B11', 'E11');
    let units = worksheet.getCell('B11');
    units.value = 'MTD';
    units.alignment = { vertical: 'middle', horizontal: 'center' };
    units.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
      color: { argb: 'FFFFFF' },
    };
    units.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '0554ef' },
      bgColor: { argb: 'FF0000FF' },
    };
    units.border = { right: { style: 'thin' } };

    worksheet.mergeCells('F11', 'G11');
    let frontgross = worksheet.getCell('F11');
    frontgross.value = 'LY MTD';
    frontgross.alignment = { vertical: 'middle', horizontal: 'center' };
    frontgross.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
      color: { argb: 'FFFFFF' },
    };
    frontgross.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '0554ef' },
      bgColor: { argb: 'FF0000FF' },
    };
    frontgross.border = { right: { style: 'thin' } };

    worksheet.mergeCells('H11', 'I11');
    let backgross = worksheet.getCell('H11');
    backgross.value = 'LM';
    backgross.alignment = { vertical: 'middle', horizontal: 'center' };
    backgross.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
      color: { argb: 'FFFFFF' },
    };
    backgross.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '0554ef' },
      bgColor: { argb: 'FF0000FF' },
    };
    backgross.border = { right: { style: 'thin' } };

    worksheet.mergeCells('J11', 'L11');
    let totalgross = worksheet.getCell('J11');
    totalgross.value = 'YTD';
    totalgross.alignment = { vertical: 'middle', horizontal: 'center' };
    totalgross.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
      color: { argb: 'FFFFFF' },
    };
    totalgross.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '0554ef' },
      bgColor: { argb: 'FF0000FF' },
    };
    totalgross.border = { right: { style: 'thin' } };
    worksheet.mergeCells('M11', 'N11');
    let LY = worksheet.getCell('M11');
    LY.value = 'LY';
    LY.alignment = { vertical: 'middle', horizontal: 'center' };
    LY.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
      color: { argb: 'FFFFFF' },
    };
    LY.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '0554ef' },
      bgColor: { argb: 'FF0000FF' },
    };
    LY.border = { right: { style: 'thin' } };
    // worksheet.addRow('');
    let Headings = [
      CurrentDate,
      CurrentDate,
      'Pace',

      'Budget',
      'Variance',
      LMYDate,
      'Variance',
      LMDate,
      'Variance',
      CurrentYear,
      'Budget	',
      'Variance',
      LastYear,
      'Variance',
    ];
    const headerRow = worksheet.addRow(Headings);
    headerRow.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
      color: { argb: 'black' },
    };
    headerRow.height = 20
    headerRow.alignment = {
      indent: 1,
      vertical: 'middle',
      horizontal: 'center',
    };
    headerRow.eachCell((cell, number) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'dbe7ff' },
        bgColor: { argb: 'FF0000FF' },
      };
      cell.border = { right: { style: 'thin' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });

    // Generate a unique file name
    FSData.forEach((d: any) => {
      const Obj = [
        d.LABLE1 == '' ? '-' : d.LABLE1 == null ? '-' : d.LABLE1,

        d.MTD == '' ? '-' : d.MTD == null ? '-' : d.MTD,
        d.PACE == '' ? '-' : d.PACE == null ? '-' : d.PACE,
        d.BUDGET == '' ? '-' : d.BUDGET == null ? '-' : d.BUDGET,
        d.VARIANCE == '' ? '-' : d.VARIANCE == null ? '-' : d.VARIANCE,
        d.LMY_MTD == '' ? '-' : d.LMY_MTD == null ? '-' : d.LMY_MTD,
        d.LMY_VARIANCE == ''
          ? '-'
          : d.LMY_VARIANCE == null
            ? '-'
            : d.LMY_VARIANCE,
        d.LM_MTD == '' ? '-' : d.LM_MTD == null ? '-' : d.LM_MTD,
        d.LM_VARIANCE == ''
          ? '-'
          : d.LM_VARIANCE == null
            ? '-'
            : d.LM_VARIANCE,
        d.YTD == '' ? '-' : d.YTD == null ? '-' : d.YTD,
        d.CY_BUDGET == ''
          ? '-'
          : d.CY_BUDGET == null
            ? '-'
            : parseInt(d.CY_BUDGET),
        d.CY_VARIANCE == ''
          ? '-'
          : d.CY_VARIANCE == null
            ? '-'
            : d.CY_VARIANCE,
        d.LY_YTD == '' ? '-' : d.LY_YTD == null ? '-' : d.LY_YTD,
        d.LY_VARIANCE == ''
          ? '-'
          : d.LY_VARIANCE == null
            ? '-'
            : d.LY_VARIANCE,
      ];
      // console.log(Obj);
      const row = worksheet.addRow(Obj);
      if (
        d.LABLE1 == 'New Units' ||
        d.LABLE1 == 'Used Units' ||
        d.LABLE1 == 'Unit Retail Sales' ||
        d.LABLE1 == 'Fleet Units' ||
        d.LABLE1 == 'Pre-Owned Units'
      ) {
        row.eachCell({ includeEmpty: true }, (cell, rowNumber) => {
          if (rowNumber > 1) {
            // Skip the header row
            cell.numFmt = '#,##0';
          }
        });
      } else if (
        d.LABLE1 == 'Selling Gross%' ||
        d.LABLE1 == 'Net to Sales' ||
        d.LABLE1 == 'Net to Gross' ||
        d.LABLE1 == 'Fixed Expenses%' ||
        d.LABLE1 == 'Other Expenses%'
      ) {
        row.eachCell({ includeEmpty: true }, (cell, rowNumber) => {
          if (rowNumber > 1) {
            if (typeof cell.value === 'number') {
              cell.value = cell.value / 100;
            }
            cell.numFmt = '0%';
          }
        });
      } else {
        row.eachCell({ includeEmpty: true }, (cell, rowNumber) => {
          if (rowNumber > 1) {
            // Skip the header row
            cell.numFmt = '$#,##0';
          }
        });
      }
      row.getCell(1).alignment = {
        indent: 1,
        vertical: 'middle',
        horizontal: 'left',
      };
      row.font = { name: 'Arial', family: 4, size: 8 };
      const lastRow = worksheet.lastRow?.number ?? 0;
      row.eachCell((cell) => {
        cell.border = {
          right: { style: 'thin' },
          bottom: row.number === lastRow ? { style: 'thin' } : undefined
        };
        cell.alignment = {
          vertical: 'top',
          horizontal: 'right',
          indent: 1
        };
      });
      if (row.number % 2) {
        row.eachCell((cell, number) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'f9fbff' },
            bgColor: { argb: 'FF0000FF' },
          };
        });
      }
      if (
        d.LABLE1 === 'Unit Retail Sales' ||
        d.LABLE1 === 'Total Sales Gross' ||
        d.LABLE1 === 'Total Store Gross' ||
        d.LABLE1 === 'Selling Gross' ||
        d.LABLE1 === 'Selling Gross%' ||
        d.LABLE1 === 'Total Expenses' ||
        d.LABLE1 === 'Operating Profit' ||
        d.LABLE1 === 'Net Adds/Deducts' ||
        d.LABLE1 === 'Net Profit' ||
        d.LABLE1 === 'Total Store Super Gross'
      ) {
        row.eachCell((cell) => {
          cell.font = { name: 'Arial', family: 4, size: 8, bold: true };
        });
      }
    });

    worksheet.getColumn(1).width = 23;
    worksheet.getColumn(1).alignment = {
      indent: 1,
      vertical: 'middle',
      horizontal: 'left',
    };
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
    worksheet.addRow([]);
    workbook.xlsx.writeBuffer().then((data: any) => {
      const blob = new Blob([data], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      FileSaver.saveAs(
        blob,
        'Financial Summary_' + DATE_EXTENSION + EXCEL_EXTENSION
      );
    });
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


