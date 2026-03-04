import { Component, HostListener, Injector } from '@angular/core';
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
import { ToastContainer } from '../../../../Layout/toast-container/toast-container';
import { IncomestatementtrendDetails } from '../incomestatementtrend-details/incomestatementtrend-details';
import { IncomestatementtrendGraph } from '../incomestatementtrend-graph/incomestatementtrend-graph';
import { Stores } from '../../../../CommonFilters/stores/stores';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [SharedModule, BsDatepickerModule, ToastContainer, Stores],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
  providers: [DatePipe]
})
export class Dashboard {
  [x: string]: any;
  Current_Date: any;

  gridshow: boolean = false;
  monthgridshow: boolean = false;
  FSData: any = [];
  StoreNamesHeadings: any = [];
  MonthsHeadings: any = [];

  NoData: boolean = false;
  fromnewdate = new Date();
  date: any = new Date(
    this.fromnewdate.setFullYear(this.fromnewdate.getFullYear())
  );
  ExpenseTrendByStoreKeys: any[] = [];
  AllDatakeys: any[] = [];
  ExpenseTrendByStore_Excel: any;
  XpenseTrendByStoreKeys: string[] = [];
  ExpenseTrendByStore: any;
  ExpenseTrendByStore_ExcelMonth: any;
  XpenseTrendByStoreKeysMonth: string[] = [];
  ExpenseTrendByStoreKeysMonth: any;
  AllDatakeysMonth: any;
  ExpenseTrendByStoreMonth: any;
  Filter: any = 'StoreSummary';
  StoreName: any;
  SubFilter: any;
  SelectedTab: any[] = [];
  SubSelectedTab1: any[] = [];
  Month: any;
  // stores: any;
  selectedstorevalues: any;

  fromdate: any;
  todate: any;
  tonewdate: any = new Date();
  solutionurl: any = environment.apiUrl;
  StoreValues: any = [

  ];
  groups: any = 1;
  selectedstorename: any = 'All Stores';
  header: any = [
    {
      type: 'Bar',
      storeIds: this.StoreValues,
      fromdate: '01' + '-' + '01' + '-' + this.date.getFullYear(),
      todate:
        ('0' + (this.tonewdate.getMonth() + 1)).slice(-2) +
        '-' +
        '01' +
        '-' +
        this.tonewdate.getFullYear(),
      groups: this.groups
    },
  ];
  popup: any = [{ type: 'Popup' }];
  reportOpenSub!: Subscription;
  reportGetting!: Subscription;
  Pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;

  StartDate: Date = new Date(new Date().setFullYear(new Date().getFullYear() - 1));
  EndDate: Date = new Date();
  StartMonth!: Date;
  EndMonth!: Date;
  bsConfig: Partial<BsDatepickerConfig> = {
    dateInputFormat: 'MMMM/YYYY',
    minMode: 'month',
    maxDate: new Date()
  };
  activePopover: number = -1;
  groupsArray: any = [];
  storename: any = ''
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  groupId: any = 0;
  storeIds: any = '0';
  stores: any = [];
  storesFilterData: any = {
    'groupsArray': this.groupsArray, 'groupId': this.groupId, 'storesArray': this.stores, 'storeids': '1', 'type': 'M', 'others': 'N',
    'groupName': this.groupName, 'storename': this.storename, storecount: null, 'storedisplayname': this.storedisplayname
  };
  constructor(
    private datepipe: DatePipe,
    public apiSrvc: Api,
    private ngbmodal: NgbModal,
    private ngbmodalActive: NgbActiveModal,
    private spinner: NgxSpinnerService,
    private title: Title,
    private cp: CurrencyPipe,
    private toast: ToastService,
    private comm: common,
    private injector: Injector,
    public shared: Sharedservice,
  ) {
    if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
      // this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.ustores.split(',')
      this.groupId = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Preferences
      this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids.split(',')
    }
    if (this.shared.common.groupsandstores.length > 0) {
      this.groupsArray = this.shared.common.groupsandstores.filter((val: any) => val.sg_id != this.shared.common.reconID);
      this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
      this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_Name : this.groupName = ''
      this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
      this.getStoresandGroupsValues()
    }
    this.tonewdate = new Date();

    let today = new Date()
    // this.fromdate = '01' + '-' + '01' + '-' + this.date.getFullYear();

    let enddate = new Date(today.setDate(today.getDate() - 1));
    console.log(enddate)
    if (enddate.getMonth() == 0) {
      this.fromdate = '01-01-' + (enddate.getFullYear() - 1);
      this.todate = '01-01-' + (enddate.getFullYear());
    } else {
      console.log(enddate.getMonth())
      this.fromdate =
        ('0' + (enddate.getMonth() + 1)).slice(-2) +
        '-01' +
        '-' +
        (enddate.getFullYear() - 1);
      console.log(enddate.getMonth(), this.fromdate)


      let today = new Date();
      if (today.getDate() < 7) {
        this.todate =
          ('0' + (enddate.getMonth())).slice(-2) +
          '-' +
          ('0' + enddate.getDate()).slice(-2) +
          '-' +
          enddate.getFullYear();
        console.log(enddate.getMonth(), this.todate)
      } else {
        this.todate =
          ('0' + (enddate.getMonth() + 1)).slice(-2) +
          '-' +
          ('0' + enddate.getDate()).slice(-2) +
          '-' +
          enddate.getFullYear();
      }
      this.EndDate = new Date(this.todate)
    }

    let newdate = new Date();

    this.title.setTitle(this.comm.titleName + '-Income Statement Trend');

    const queryString = window.location.search;
    const urlParams = new URLSearchParams(queryString);

    if (localStorage.getItem('Fav') != 'Y') {
      const data = {
        title: 'Income Statement Trend',
        path1: '',
        path2: '',
        path3: '',
        Month: this.date,
        filter: this.Filter,
        stores: this.storeIds,
        groups: this.groups,
        selectedstorename: this.selectedstorename,
        fromdate:
          this.fromdate,
        todate:
          this.todate,
        count: 0,
      };

      this.apiSrvc.SetHeaderData({
        obj: data,
      });
      this.header = [
        {
          type: 'Bar',
          storeIds: this.StoreValues,
          fromdate:
            this.fromdate,
          todate:
            this.todate,
          groups: this.groups
        },
      ];
      this.StartMonth = this.StartDate;
      this.EndMonth = this.EndDate;
      this.GetDataByMonths(this.StartMonth, this.EndMonth);
    } else {
      this.getFavReports();
    }
  }

  ngOnInit(): void {

    this.getStores();
  }

  Scrollpercent: any = 0;
  updateVerticalScroll(event: any): void {
    const scrollDemo = <HTMLInputElement>document.querySelector('#scrollcent');
    this.Scrollpercent = Math.round(
      (event.target.scrollTop /
        (event.target.scrollHeight - scrollDemo.clientHeight)) *
      100
    );
  }

  applyDateChange() {
    if (!this.storeIds || this.storeIds.length === 0) {
      this.toast.show(
        'Please Select Atleast One Store',
        'warning',
        'Warning'
      );
      return;
    }

    if (!this.StartDate || !this.EndDate) {
      this.toast.show(
        'Please Enter Valid Date',
        'warning',
        'Warning'
      );
      return;
    }

    const start = new Date(this.StartDate);
    const end = new Date(this.EndDate);

    if (isNaN(start.getTime()) || isNaN(end.getTime())) {
      this.toast.show(
        'Please Enter Valid Date',
        'warning',
        'Warning'
      );
      return;
    }

    if (start > end) {
      this.toast.show(
        'Please Enter Valid Date',
        'warning',
        'Warning'
      );
      return;
    }
    const monthDiff =
      (end.getFullYear() - start.getFullYear()) * 12 +
      (end.getMonth() - start.getMonth());

    if (monthDiff < 3) {
      this.toast.show(
        'Please Select Atleast 3 Months Range',
        'warning',
        'Warning'
      );
      return;
    }

    this.StartMonth = this.StartDate;
    this.EndMonth = this.EndDate;

    this.GetDataByMonths(this.StartMonth, this.EndMonth);
  }


  FromDate: any;
  ToDate: any;
  GetDataByMonths(StartMonth: any, EndMonth: any) {
    this.spinner.show();
    this.ExpenseTrendByStoreKeysMonth = [];
    this.AllDatakeysMonth = [];
    this.ExpenseTrendByStoreMonth = [];
    this.FromDate = this.datepipe.transform(StartMonth, 'dd-MMM-yyyy');
    this.ToDate = this.datepipe.transform(EndMonth, 'dd-MMM-yyyy');
    const obj = {
      Stores: this.storeIds.toString(),
      SalesDATE: this.datepipe.transform(StartMonth, 'dd-MMM-yyyy'),
      EndDate: this.datepipe.transform(EndMonth, 'dd-MMM-yyyy'),
      UserID: 0,
    };

    let startFrom = new Date().getTime();
    const curl = environment.apiUrl + this.comm.routeEndpoint + 'GetIncomeStatementTrend';
    this.apiSrvc.postmethod(this.comm.routeEndpoint + 'GetIncomeStatementTrend', obj).subscribe(
      (x) => {
        const currentTitle = document.title;
        this.apiSrvc.logSaving(curl, {}, '', x.message, currentTitle);
        if (x.status == 200) {
          if (x.response != undefined) {
            if (x.response.length > 0) {
              this.spinner.hide();

              let resTime = (new Date().getTime() - startFrom) / 1000;
              this.monthgridshow = true;

              this.XpenseTrendByStoreKeysMonth = Object.keys(x.response[0]);
              let ETByStoreKeys_sortedMonth =
                this.XpenseTrendByStoreKeysMonth.filter((x) => {
                  return x != 'SNo' && x != 'NgClass';
                });
              let AllStore_LabelMonth = ETByStoreKeys_sortedMonth.find(
                (x: any) => {
                  if (x.toUpperCase() == 'All Stores'.toUpperCase()) return x;
                }
              );
              ETByStoreKeys_sortedMonth = ETByStoreKeys_sortedMonth.splice(2);
              this.MonthsHeadings = ETByStoreKeys_sortedMonth;
              const Cmtindex = ETByStoreKeys_sortedMonth.findIndex(
                (i) => i == 'CommentsStatus'
              );
              if (Cmtindex >= 0) {
                ETByStoreKeys_sortedMonth.splice(Cmtindex, 1);
              }
              for (var i = 0; i < ETByStoreKeys_sortedMonth.length; i++) {
                this.ExpenseTrendByStoreKeysMonth.push(
                  ETByStoreKeys_sortedMonth[i]
                );
                this.AllDatakeysMonth.push(ETByStoreKeys_sortedMonth[i]);
              }
              this.ExpenseTrendByStoreKeysMonth.push(AllStore_LabelMonth);
              this.AllDatakeysMonth.push(AllStore_LabelMonth);
              let XpenseTrendByStoreDataMonth = x.response;
              let ETByStoreData_sorted = XpenseTrendByStoreDataMonth;
              let AllStoreLabel_Data = '';
              this.ExpenseTrendByStoreMonth = x.response;

              if (this.ExpenseTrendByStoreKeysMonth.length == 0) {
                this.NoData = true;
              } else {
                this.NoData = false;
              }
              this.spinner.hide();
            } else {
              this.spinner.hide();
              this.NoData = true;
            }
          } else {
            this.ExpenseTrendByStoreMonth = [];
            this.AllDatakeysMonth = [];
            this.spinner.hide();
            this.NoData = true;
          }
        } else {
          this.toast.show(
            x.status,
            'danger',
            'Error'
          );

          this.spinner.hide();
          this.NoData = true;
          let resTime = (new Date().getTime() - startFrom) / 1000;
        }
      },
      (error) => {
        this.toast.show(
          '502 Bad Gate Way Error',
          'danger',
          'Error'
        );
        this.spinner.hide();
        this.NoData = true;
      }
    );
  }

  isUnitRow(label: string): boolean {
    return ['New Units', 'Pre-Owned Units', 'Unit Retail Sales', 'Fleet Units', 'Variable Expenses%',
      'Personnel Expenses%', 'Semi-Fixed Expenses%', 'Fixed Expenses%', 'Other Expenses%'].includes(label);
  }
  isNonClickable(value: any, itemKey: string, label: string): boolean {
    return value === null || value === undefined || value === '' ||
      itemKey === 'YTD' || itemKey === 'LYTD' || this.isUnitRow(label);
  }
  isBoldRow(label: string): boolean {
    return [
      'Unit Retail Sales',
      'Pure Gross',
      'Variable Gross',
      'Total Fixed',
      'Total Store Gross',
      'Total Expenses',
      'Operating Profit',
      'Net Adds/Deducts',
      'Net Profit'
    ].includes(label);
  }

  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    }
    return false;
  }

  SFstate: any;

  ngAfterViewInit(): void {
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.shared.common.pageName == 'Income Statement Trend') {
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
      // //console.log(res);
      if (this.reportOpenSub != undefined) {
        if (res.obj.Module == 'Income Statement Trend') {
          document.getElementById('report')?.click();
        }
      }
    });
    this.excel = this.apiSrvc.getExportToExcelAllReports().subscribe((res: { obj: { state: boolean; title: string; }; }) => {
      if (this.excel != undefined) {
        this.SFstate = res.obj.state;
        if (res.obj.title == 'Income Statement Trend') {
          if (res.obj.state == true) {
            this.exportAsXLSX();
          }
        }
      }
    });
    this.reportGetting = this.apiSrvc.GetReports().subscribe((data) => {
      ////console.log(data)
      if (this.reportGetting != undefined) {
        if (data.obj.Reference == 'Income Statement Trend') {
          if (data.obj.header == undefined) {
            this.date = data.obj.month;
            this.Month = data.obj.month;
            // this.Month =  ('0' +( new Date(data.obj.month).getMonth())).slice(-2)+'-01'+'-'+new Date (data.obj.month).getFullYear()
            this.StoreValues = data.obj.storeValues;
            this.selectedstorename = data.obj.selectedstorename;
            this.StoreName = data.obj.Sname;
            this.Filter = data.obj.filters;
            this.SubFilter = data.obj.subfilters;
            this.groups = data.obj.groups;
            let newdate = new Date();
            (this.fromdate = this.datepipe.transform(
              data.obj.FromDate,

              'dd-MMM-yyyy'
            )),
              (this.todate = this.datepipe.transform(
                data.obj.ToDate,

                'dd-MMM-yyyy'
              ));
          } else {
            if (data.obj.header == 'Yes') {
              this.StoreValues = data.obj.storeValues;
              //console.log(this.StoreValues);
              if (this.StoreValues != '') {
                let storename = this.comm.groupsandstores.filter(
                  (v: any) => v.sg_id == this.groups
                )[0].Stores;
                this.selectedstorename = storename.filter(
                  (v: any) => v.ID == this.StoreValues
                )[0].storename;
              }

              //console.log(this.selectedstorename);
            }
          }

          // this.DataSelection(this.Filter);
          const headerdata = {
            title: 'Income Statement Trend',
            path1: '',
            path2: '',
            path3: '',
            Month: this.Month,
            filter: this.Filter,
            stores: this.StoreValues,
            fromdate: this.fromdate,
            todate: this.todate,
            groups: this.groups,
            selectedstorename: this.selectedstorename,
          };
          this.apiSrvc.SetHeaderData({
            obj: headerdata,
          });
          this.header = [
            {
              type: 'Bar',
              storeIds: this.StoreValues,
              fromdate: this.fromdate,
              todate: this.todate,
              groups: this.groups
            },
          ];
          // if (this.Filter == 'StoreSummary') {
          // this.GetData();
          // } else {
          this.GetDataByMonths(this.StartMonth, this.EndMonth);
          // }
        }
      }
    });

    this.print = this.apiSrvc.getExportToPrintAllReports().subscribe((res: { obj: { state: boolean; title: string; }; }) => {
      if (this.print != undefined) {
        if (res.obj.title == 'Income Statement Trend') {
          if (res.obj.state == true) {
            this.GetPrintData();
          }
        }
      }
    });

    this.Pdf = this.apiSrvc.getExportToPDFAllReports().subscribe((res: { obj: { state: boolean; title: string; }; }) => {
      if (this.Pdf != undefined) {
        if (res.obj.title == 'Income Statement Trend') {
          if (res.obj.state == true) {
            this.generatePDF();
          }
        }
      }
    });
    this.email = this.apiSrvc.getExportToEmailPDFAllReports().subscribe((res: { obj: { state: boolean; title: string; }; }) => {

      if (this.email != undefined) {
        if (res.obj.title == 'Income Statement Trend') {
          if (res.obj.state == true) {
            // this.sendEmailData(res.obj.Email, res.obj.notes, res.obj.from);
          }
        }
      }

    });
  }
  ngOnDestroy() {
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
  getStores() {
    // this.stores = environment.stores;
    this.selectedstorevalues = [];
  }

  openDetails(Object: any, monthname: any, ref: any, item: any) {
    console.log(Object, monthname, ref, item, this.selectedstorename);
    const date = '01-' + monthname;
    const myDate = this.datepipe.transform(new Date(date), 'dd-MMM-yyyy');
    const DetailsSF = this.ngbmodal.open(IncomestatementtrendDetails, {
      size: 'xl',
      backdrop: 'static',
      injector: Injector.create({
        providers: [
          { provide: CurrencyPipe, useClass: CurrencyPipe }
        ],
        parent: this.injector
      })
    });
    DetailsSF.componentInstance.Fsdetails = {
      TYPE: Object.LABLEVAL,
      NAME: Object.LABLE,
      STORES: this.storeIds.toString(),
      LatestDate: myDate,
      STORENAME: 'WesternAuto',
      Group: this.groups,
    };
  }

  openGraph(Object: any, monthname: any, ref: any, item: any) {
    const DetailsSF = this.ngbmodal.open(IncomestatementtrendGraph, {
      size: 'xl',
      backdrop: 'static',
      injector: Injector.create({
        providers: [
          { provide: CurrencyPipe, useClass: CurrencyPipe }
        ],
        parent: this.injector
      })
    });
    DetailsSF.componentInstance.INSTGRAPHdetails = {
      ITEM: item,
      TYPE: item.LABLEVAL,
      NAME: item.LABLE,
      STORES: this.storeIds.toString(),
      STORENAME: 'WesternAuto',
    };
  }

  Favreports: any = [];
  getFavReports() {
    const obj = {
      Id: localStorage.getItem('Fav_id'),
      expression: '',
    };
    this.apiSrvc.postmethod('favouritereports/get', obj).subscribe((res) => {
      if (res.status == 200) {
        if (res.response.length > 0) {
          this.Favreports = res.response;
          let dates = this.Favreports[0].Fav_Report_Value.split(',');
          let startdate = new Date(dates[0]);
          let enddate = new Date(dates[1]);
          // //console.log(startdate, enddate);
          (this.fromdate = this.datepipe.transform(startdate, 'dd-MMM-yyyy')),
            (this.todate = this.datepipe.transform(enddate, 'dd-MMM-yyyy'));
          this.StoreValues = res.response[1].Fav_Report_Value;
          // this.DataSelection(this.Filter);
          this.GetDataByMonths(this.StartMonth, this.EndMonth);
          localStorage.setItem('Fav', 'N');
          const data = {
            title: 'Income Statement Trend',
            path1: '',
            path2: '',
            path3: '',
            Month: this.date,
            filter: this.Filter,
            stores: this.StoreValues,
            fromdate: this.datepipe.transform(startdate, 'MM-dd-yyyy'),
            todate:
              ('0' + (enddate.getMonth() + 1)).slice(-2) +
              '-' +
              '01' +
              '-' +
              enddate.getFullYear(),
          };
          // // //console.log(data)
          this.apiSrvc.SetHeaderData({ obj: data });
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
  ExcelStoreNames: any = [];
  exportAsXLSX(): void {
    let ISmonthByData = this.ExpenseTrendByStoreMonth.map(
      (_arrayElement: any) => Object.assign({}, _arrayElement)
    );
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet('Income Statement Trend');
    worksheet.views = [
      {
        state: 'frozen',
        ySplit: 13, // Number of rows to freeze (2 means the first two rows are frozen)
        topLeftCell: 'A14', // Specify the cell to start freezing from (in this case, the third row)
        showGridLines: false,
      },
    ];
    worksheet.addRow('');
    const titleRow = worksheet.addRow(['Income Statement Trend']);
    titleRow.eachCell((cell, number) => {
      cell.alignment = { indent: 1, vertical: 'top', horizontal: 'left' };
    });
    titleRow.font = { name: 'Arial', family: 4, size: 12, bold: true };
    titleRow.worksheet.mergeCells('A2', 'D2');

    worksheet.addRow('');
    const DateToday = this.datepipe.transform(
      new Date(),
      'MM.dd.yyyy h:mm:ss a'
    );
    const DATE_EXTENSION = this.datepipe.transform(new Date(), 'MMddyyyy');
    worksheet.addRow([DateToday]).font = { name: 'Arial', family: 4, size: 9 };

    const StartDate: any = this.datepipe.transform(this.StartDate, 'MMM-yy');
    const EndDate: any = this.datepipe.transform(this.EndDate, 'MMM-yy');
    const Header = [EndDate + ' to ' + StartDate];
    for (let i = 0; i < this.MonthsHeadings.length; i++) {
      Header.push(this.MonthsHeadings[i]);
    }
    const ReportFilter = worksheet.addRow(['Report Filters :']);
    ReportFilter.font = { name: 'Arial', family: 4, size: 10, bold: true };

    const SummaryType = worksheet.addRow(['Summary Type :']);
    const summarytype = worksheet.getCell('B6');
    summarytype.value = 'Month Summary';
    summarytype.font = { name: 'Arial', family: 4, size: 9 };
    summarytype.alignment = { vertical: 'top', horizontal: 'left' };
    SummaryType.getCell(1).font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
    };
    const DateMonth = worksheet.addRow(['Time Frame :']);
    const datemonth = worksheet.getCell('B7');
    datemonth.value = EndDate + ' to ' + StartDate;
    datemonth.font = { name: 'Arial', family: 4, size: 9 };
    datemonth.alignment = { vertical: 'top', horizontal: 'left' };
    DateMonth.getCell(1).font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
    };
    // const Groups = worksheet.getCell('A8');
    // Groups.value = 'Group :';
    // Groups.font = { name: 'Arial', family: 4, size: 9, bold: true };
    // const groups = worksheet.getCell('B8');
    // groups.value =
    //   this.groups = 'WesternAuto';
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
    worksheet.mergeCells('B9', 'K11');
    const Stores = worksheet.getCell('A9');
    Stores.value = 'Store :';
    const stores = worksheet.getCell('B9');
    stores.value = storeValue;
    stores.font = { name: 'Arial', family: 4, size: 9 };
    stores.alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    Stores.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
    };
    worksheet.addRow('');
    const headerRow = worksheet.addRow(Header);
    // //console.log(headerRow);
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
    headerRow.eachCell((cell, number) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '2a91f0' },
        bgColor: { argb: 'FF0000FF' },
      };
      cell.border = { right: { style: 'thin' } };
      cell.alignment = { vertical: 'top', horizontal: 'center' };
    });
    ISmonthByData.forEach((d: any) => {
      var Obj = [d.LABLE];
      this.MonthsHeadings.forEach((e: any) => {
        Obj = [...Obj, d[e] == '' ? '-' : d[e] == null ? '-' : parseInt(d[e])];
      });
      // //console.log(Obj);
      const row = worksheet.addRow(Obj);
      row.getCell(1).alignment = {
        indent: 1,
        vertical: 'middle',
        horizontal: 'left',
      };
      if (row.number % 2) {
        row.eachCell((cell, number) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'e5e5e5' },
            bgColor: { argb: 'FF0000FF' },
          };
        });
      }
      row.font = { name: 'Arial', family: 4, size: 8 };
      if (this.isBoldRow(d.LABLE)) {
        row.eachCell((cell) => {
          cell.font = { bold: true, name: 'Arial', family: 4, size: 8 };
        });
      }
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
      if (
        d.LABLE == 'New Units' ||
        d.LABLE == 'Pre-Owned Units' ||
        d.LABLE == 'Unit Retail Sales' ||
        d.LABLE == 'Wholesale Units' ||
        d.LABLE == 'Fleet Units'
      ) {
        row.eachCell({ includeEmpty: true }, (cell, rowNumber) => {
          if (rowNumber > 1) {
            // Skip the header row
            cell.numFmt = '#,##0';
          }
        });
      } else if (
        d.LABLE == 'Variable Expenses%' ||
        d.LABLE == 'Personnel Expenses%' ||
        d.LABLE == 'Semi-Fixed Expenses%' ||
        d.LABLE == 'Fixed Expenses%' ||
        d.LABLE == 'Other Expenses%'
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
    });
    worksheet.getColumn(1).width = 30;
    worksheet.getColumn(1).alignment = {
      indent: 1,
      vertical: 'top',
      horizontal: 'left',
    };
    worksheet.getColumn(2).width = 13;
    worksheet.getColumn(3).width = 13;
    worksheet.getColumn(4).width = 13;
    worksheet.getColumn(5).width = 13;
    worksheet.getColumn(6).width = 13;
    worksheet.getColumn(7).width = 13;
    worksheet.getColumn(8).width = 13;
    worksheet.getColumn(9).width = 13;
    worksheet.getColumn(10).width = 13;
    worksheet.getColumn(11).width = 13;
    worksheet.getColumn(12).width = 13;
    worksheet.getColumn(13).width = 13;
    worksheet.getColumn(14).width = 13;
    worksheet.getColumn(15).width = 13;
    worksheet.getColumn(16).width = 13;
    worksheet.addRow([]);
    workbook.xlsx.writeBuffer().then((data: any) => {
      const blob = new Blob([data], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      FileSaver.saveAs(
        blob,
        'Income Statement Trend_' + DATE_EXTENSION + EXCEL_EXTENSION
      );
    });
    // });
  }

  GetPrintData() {
    window.print();
  }

  generatePDF() {
    this.spinner.show();
    const printContents = document.getElementById(
      'IncomeStatementTrend'
    )!.innerHTML;
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
    <title>Income Statement Trend</title>
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
                .store {
                  background-color: #e5e5e5 !important;
                  color: black !important;
                 }
                 .setPerens:after {
                  content: ")";
                 }
                 .setPerens:before {
                  content: "(";
                 }
                 .negative {
                  color: red;
                 }
                 .notbold {
                  font-family: "Roboto";
                  font-size: 0.8rem;
                  white-space: nowrap;
                 }
                 .boldlabel {
                  font-family: "RobotoBold";
                  color: #2a91f0 !important;
                  font-size: 0.8rem;
                  white-space: nowrap;
                 }
                .performance-scorecard .table>:not(:first-child) {
                  border-top: 0px solid #ffa51a
                  }
                  .performance-scorecard .table {
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
                  .performance-scorecard .table th:first-child {
                    position: sticky;
                    left: 0;
                    z-index: 1;
                    // background-color: #337ab7;
                   }
                   .performance-scorecard .table td:first-child {
                    position: sticky;
                    left: 0;
                    z-index: 1;
                    // background-color: #337ab7;
                   }
                   .performance-scorecard .table tr:nth-child(odd) {
                    // background-color: #e9ecef;
                    background-color: #ffffff;
            
                   }
                   .performance-scorecard .table tr:nth-child(even) {
                    background-color: #ffffff;
                   }
                   .performance-scorecard .table .spacer {
                    // width: 50px !important;
                    background-color: #cfd6de !important;
                    border-left: 1px solid #cfd6de !important;
                    border-bottom: 1px solid #cfd6de !important;
                    border-top: 1px solid #cfd6de !important;
                  }
                  .performance-scorecard .table .hidden {
                    display: none !important;
                   }
                   .performance-scorecard .table .bdr-rt {
                    border-right: 1px solid #abd0ec;
                   }
                   .performance-scorecard .table thead {
                    position: sticky;
                    top: 0;
                    z-index: 99;
                    font-family: 'FaktPro-Bold';
                    font-size: 0.8rem;
                   }
                   .performance-scorecard .table thead th {
                    padding: 5px 10px;
                    margin: 0px;
                   }
                   .performance-scorecard .table thead  .bdr-btm {
                    border-bottom: #005fa3;
                    }
                    .performance-scorecard .table thead  tr:nth-child(1) {
                      background-color: #fff !important;
                      color: #000;
                      // color: #fff;
                      text-transform: uppercase;
                      border-bottom: #cfd6de;
                     }
                     .performance-scorecard .table thead tr:nth-child(2) {
                      background-color: #337ab7 !important;
                      color: #fff;
                      text-transform: uppercase;
                      border-bottom: #cfd6de;
                      box-shadow: inset 0 1px 0 0 #cfd6de;
                    }
                    .performance-scorecard .table thead  tr:nth-child(3) {
            
                      background-color: #fff !important;
                      color: #000;
                      text-transform: uppercase;
                      border-bottom: #cfd6de;
                      box-shadow: inset 0 1px 0 0 #cfd6de;
                    }
                    .performance-scorecard .table thead tr:nth-child(3) th :nth-child(1) {
                      background-color: #337ab7 !important;
                      color: #fff;
                     }
                     .performance-scorecard .table tbody {
                      font-family: 'FaktPro-Normal';
                      font-size: .9rem;
                     }
                     .performance-scorecard .table tbody td {
                      padding: 2px 10px;
                      margin: 0px;
                      border: 1px solid #cfd6de
                      }
                      .performance-scorecard .table tbody  tr {
                        border-bottom: 1px solid #37a6f8;
                        border-left: 1px solid #37a6f8
                      }
                      .performance-scorecard .table tbody td:first-child {
                        text-align: start;
                        box-shadow: inset -1px 0 0 0 #cfd6de;
                      }
                      .performance-scorecard .table tbody  tr td:not(:first-child) {
                        text-align: right !important;
            
                      } 
                      .performance-scorecard .table tbody .sub-title {
                       font-size: .8rem !important;
                      }
                      .performance-scorecard .table tbody .sub-subtitle{
                      font-size: .7rem !important;
                      }
                     .performance-scorecard .table tbody  td:nth-child(2){ padding: 2px 10px;
                      margin: 0px;
                     }
                     .performance-scorecard .table tbody .text-bold {
                      font-family: 'FaktPro-Bold';
                      }
                     .performance-scorecard .table tbody .darkred-bg {
                      background-color: #282828 !important;
                      color: #fff;
                     }
                     .performance-scorecard .table tbody .lightblue-bg {
                      background-color: #646e7a !important;
                      color: #fff;
                     }
                     .performance-scorecard .table tbody .gold-bg {
                      background-color: #ffa51a;
                      color: #fff;
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
        let imgWidth = 285;
        let pageHeight = 227.2;
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
        doc.save('Income Statement Trend.pdf');
        // popupWin.close();
        this.spinner.hide();
      });
  }

  sendEmailData(Email: any, notes: any, from: any) {
    this.spinner.show();
    const printContents = document.getElementById(
      'IncomeStatementTrend'
    )!.innerHTML;
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
            <title>Income Statement Trend</title>
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
                .store {
                  background-color: #e5e5e5 !important;
                  color: black !important;
                 }
                 .setPerens:after {
                  content: ")";
                 }
                 .setPerens:before {
                  content: "(";
                 }
                 .negative {
                  color: red;
                 }
                 .notbold {
                  font-family: "Roboto";
                  font-size: 0.8rem;
                  white-space: nowrap;
                 }
                 .boldlabel {
                  font-family: "RobotoBold";
                  color: #2a91f0 !important;
                  font-size: 0.8rem;
                  white-space: nowrap;
                 }
                .performance-scorecard .table>:not(:first-child) {
                  border-top: 0px solid #ffa51a
                  }
                  .performance-scorecard .table {
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
                  .performance-scorecard .table th:first-child {
                    position: sticky;
                    left: 0;
                    z-index: 1;
                    // background-color: #337ab7;
                   }
                   .performance-scorecard .table td:first-child {
                    position: sticky;
                    left: 0;
                    z-index: 1;
                    // background-color: #337ab7;
                   }
                   .performance-scorecard .table tr:nth-child(odd) {
                    // background-color: #e9ecef;
                    background-color: #ffffff;
            
                   }
                   .performance-scorecard .table tr:nth-child(even) {
                    background-color: #ffffff;
                   }
                   .performance-scorecard .table .spacer {
                    // width: 50px !important;
                    background-color: #cfd6de !important;
                    border-left: 1px solid #cfd6de !important;
                    border-bottom: 1px solid #cfd6de !important;
                    border-top: 1px solid #cfd6de !important;
                  }
                  .performance-scorecard .table .hidden {
                    display: none !important;
                   }
                   .performance-scorecard .table .bdr-rt {
                    border-right: 1px solid #abd0ec;
                   }
                   .performance-scorecard .table thead {
                    position: sticky;
                    top: 0;
                    z-index: 99;
                    font-family: 'FaktPro-Bold';
                    font-size: 0.8rem;
                   }
                   .performance-scorecard .table thead th {
                    padding: 5px 10px;
                    margin: 0px;
                   }
                   .performance-scorecard .table thead  .bdr-btm {
                    border-bottom: #005fa3;
                    }
                    .performance-scorecard .table thead  tr:nth-child(1) {
                      background-color: #fff !important;
                      color: #000;
                      // color: #fff;
                      text-transform: uppercase;
                      border-bottom: #cfd6de;
                     }
                     .performance-scorecard .table thead tr:nth-child(2) {
                      background-color: #337ab7 !important;
                      color: #fff;
                      text-transform: uppercase;
                      border-bottom: #cfd6de;
                      box-shadow: inset 0 1px 0 0 #cfd6de;
                    }
                    .performance-scorecard .table thead  tr:nth-child(3) {
            
                      background-color: #fff !important;
                      color: #000;
                      text-transform: uppercase;
                      border-bottom: #cfd6de;
                      box-shadow: inset 0 1px 0 0 #cfd6de;
                    }
                    .performance-scorecard .table thead tr:nth-child(3) th :nth-child(1) {
                      background-color: #337ab7 !important;
                      color: #fff;
                     }
                     .performance-scorecard .table tbody {
                      font-family: 'FaktPro-Normal';
                      font-size: .9rem;
                     }
                     .performance-scorecard .table tbody td {
                      padding: 2px 10px;
                      margin: 0px;
                      border: 1px solid #cfd6de
                      }
                      .performance-scorecard .table tbody  tr {
                        border-bottom: 1px solid #37a6f8;
                        border-left: 1px solid #37a6f8
                      }
                      .performance-scorecard .table tbody td:first-child {
                        text-align: start;
                        box-shadow: inset -1px 0 0 0 #cfd6de;
                      }
                      .performance-scorecard .table tbody  tr td:not(:first-child) {
                        text-align: right !important;
            
                      } 
                      .performance-scorecard .table tbody .sub-title {
                       font-size: .8rem !important;
                      }
                      .performance-scorecard .table tbody .sub-subtitle{
                      font-size: .7rem !important;
                      }
                     .performance-scorecard .table tbody  td:nth-child(2){ padding: 2px 10px;
                      margin: 0px;
                     }
                     .performance-scorecard .table tbody .text-bold {
                      font-family: 'FaktPro-Bold';
                      }
                     .performance-scorecard .table tbody .darkred-bg {
                      background-color: #282828 !important;
                      color: #fff;
                     }
                     .performance-scorecard .table tbody .lightblue-bg {
                      background-color: #646e7a !important;
                      color: #fff;
                     }
                     .performance-scorecard .table tbody .gold-bg {
                      background-color: #ffa51a;
                      color: #fff;
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
        let imgWidth = 285;
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
        const pdfFile = this.blobToFile(pdfBlob, 'Income Statement Trend.pdf');
        const formData = new FormData();
        formData.append('to_email', Email);
        formData.append('subject', 'Income Statement Trend');
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
}

