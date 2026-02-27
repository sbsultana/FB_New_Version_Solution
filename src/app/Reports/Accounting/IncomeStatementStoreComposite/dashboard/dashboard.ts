import { Component, Injector } from '@angular/core';
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
import { IncomestatementstoreDetails } from '../incomestatementstore-details/incomestatementstore-details';
import { ToastContainer } from '../../../../Layout/toast-container/toast-container';

@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [SharedModule, BsDatepickerModule, ToastContainer],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss'
})
export class Dashboard {
  Current_Date: any;
  gridshow: boolean = false;
  monthgridshow: boolean = false;
  FSData: any = [];
  StoreNamesHeadings: any = [];
  MonthsHeadings: any = [];
  NoData: boolean = false;
  ExpenseTrendByStoreKeys: any = [];
  AllDatakeys: any = [];
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
  SelectedTab: any = [];
  SubSelectedTab1: any = [];
  Month: any;
  stores: any;
  selectedstorevalues: any;
  selectedstorename: any;
  StoreValues: any = '0';
  popup: any = [{ type: 'Popup' }];
  groups: any = 1;
  gridvisibility: any;
  newdate: Date = new Date();
  ShowHideBudget: any = 'Hide';
  date: any = new Date();
  header: any = [{ type: 'Bar', storeIds: this.StoreValues, Month: this.date, groups: this.groups, ShowHideBudget: this.ShowHideBudget }]
  selectedDate: Date = new Date();
  currentMonth!: Date;
  constructor(
    private datepipe: DatePipe,
    public apiSrvc: Api,
    private ngbmodal: NgbModal,
    private ngbmodalActive: NgbActiveModal,
    private spinner: NgxSpinnerService,
    private title: Title,
    private toast: ToastService,
    private comm: common,
    private injector: Injector
  ) {
    if (localStorage.getItem('UserDetails') != null) {
      this.StoreValues = JSON.parse(
        localStorage.getItem('UserDetails')!
      ).Store_Ids;

      if (this.StoreValues.toString().indexOf(',') > 0) {
        this.gridvisibility = 'DL';
      } else {
        this.gridvisibility = 'SL';
      }
    }
    const lastMonth = new Date();
    let today = new Date();
    if (today.getDate() < 5) {
      this.date = new Date(lastMonth.setMonth(lastMonth.getMonth() - 1));
    } else {
      this.date = new Date(lastMonth.setMonth(lastMonth.getMonth()));
    }

    this.Month =
      ('0' + (this.date.getMonth() + 1)).slice(-2) +
      '-' +
      ('0' + this.date.getDate()).slice(-2) +
      '-' +
      this.date.getFullYear();
    this.title.setTitle(this.comm.titleName + '-Income Statement Store');
    if (localStorage.getItem('Fav') != 'Y') {
      const data = {
        title: 'Income Statement Store',
        path1: '',
        path2: '',
        path3: '',
        Month: this.date,
        // stores: 0,
        filter: this.Filter,
        stores: 2,
        ShowHideBudget: this.ShowHideBudget,
        groups: this.groups,
        count: 0
      };
      this.apiSrvc.SetHeaderData({
        obj: data,
      });
      let date = ('0' + (this.date.getMonth() + 1)).slice(-2) +
        '-01' +
        '-' +
        this.date.getFullYear();

      this.header = [{
        type: 'Bar', storeIds: this.StoreValues, Month: this.date, groups: this.groups,
        ShowHideBudget: this.ShowHideBudget,
      }]
      this.currentMonth = this.selectedDate;
      this.GetData(this.currentMonth);
    } else {
    }
    window.addEventListener('scroll', function () {
      const maxHeight = document.body.scrollHeight - window.innerHeight;
    });
  }
  ngOnInit(): void {
    this.getStores();

  }
  Scrollpercent: any = 0;
  updateVerticalScroll(event: any): void {
    const scrollDemo = document.querySelector('#scrollcent') as HTMLElement;
    this.Scrollpercent = Math.round(
      (event.target.scrollTop /
        (event.target.scrollHeight - scrollDemo.clientHeight)) *
      100
    );

  }
  sanitizeValue(value: any): string {
    if (value == null) return '-';
    return value.toString().replace(/-/g, '');
  }

  formatKey(key: string): string {
    const upperKey = key.toUpperCase();

    if (upperKey === 'BDGT_TOTAL') {
      return 'BUDGET TOTAL';
    } else if (upperKey === 'VARBDGT') {
      return 'BUDGET VARIANCE';
    } else if (key.toLowerCase().startsWith('bdgt_')) {
      return 'BUDGET';
    } else if (key.toLowerCase().startsWith('var')) {
      return 'VARIANCE';
    }

    return key;
  }


  excelformatKey(key: string): string {
    const upperKey = key.toUpperCase();

    if (upperKey === 'BDGT_TOTAL') {
      return 'BUDGET TOTAL';
    } else if (upperKey === 'VARBDGT') {
      return 'BUDGET VARIANCE';
    } else if (key.toLowerCase().startsWith('bdgt_')) {
      return 'BUDGET';
    } else if (key.toLowerCase().startsWith('var')) {
      return 'VARIANCE';
    }

    return key;
  }
  isKeyVisible(key: string): boolean {
    return key !== '-1';
  }

  bsConfig: Partial<BsDatepickerConfig> = {
    dateInputFormat: 'MMMM/YYYY',
    minMode: 'month',
    maxDate: new Date()
  };


  applyDateChange() {
    this.currentMonth = this.selectedDate;
    this.GetData(this.currentMonth);
  }
  GetData(date: any) {
    ////console.log(this.Month)
    this.spinner.show();
    this.ExpenseTrendByStoreKeys = [];
    this.AllDatakeys = [];
    const DateToday = this.datepipe.transform(
      new Date(date),
      'dd-MMM-yyyy'
    );
    const obj = {
      SalesDate: DateToday,
      Stores: 2,
      UserID: 0,
    };
    let startFrom = new Date().getTime();
    const curl = environment.apiUrl + 'WesternAuto/GetIncomeStatementStoreComposite';
    this.apiSrvc
      .postmethod(this.comm.routeEndpoint + 'GetIncomeStatementStoreComposite', obj)
      .subscribe(
        (x) => {
          const currentTitle = document.title;
          this.apiSrvc.logSaving(curl, {}, '', x.message, currentTitle);
          if (x.status == 200) {
            if (x.response != undefined) {
              if (x.response.length > 0) {
                let resTime = (new Date().getTime() - startFrom) / 1000;
                this.gridshow = true;
                this.XpenseTrendByStoreKeys = Object.keys(x.response[0]);
                let ETByStoreKeys_sorted = this.XpenseTrendByStoreKeys.filter((x) => {
                  if (x === 'SNo' || x === 'NgClass') return false;

                  if (this.ShowHideBudget === 'Show') {
                    return true; // include all keys
                  }
                  const isDynamicBudget = /^BDGT_\d+$/.test(x);
                  const isDynamicVariance = /^VAR\d+$/.test(x);
                  const isBudgetTotal = x === 'BDGT_TOTAL';
                  const isBudgetVariance = x === 'VARBDGT';
                  return !isDynamicBudget && !isDynamicVariance && !isBudgetTotal && !isBudgetVariance;
                });
                console.log('Filtered keys:', ETByStoreKeys_sorted);
                console.log('Show/Hide Budget setting:', this.ShowHideBudget);
                console.log('ETByStoreKeys_sorted', ETByStoreKeys_sorted)


                let AllStore_Label = ETByStoreKeys_sorted.find((x: any) => {
                  if (x.toUpperCase() == 'All Stores'.toUpperCase()) return x;
                });
                console.log('AllStore_Label', AllStore_Label)
                const lable = ETByStoreKeys_sorted.findIndex(
                  (i) => i == 'LABLE'
                );
                ETByStoreKeys_sorted.splice(lable, 1);
                ETByStoreKeys_sorted = ETByStoreKeys_sorted.splice(1);
                this.StoreNamesHeadings = ETByStoreKeys_sorted;
                const index = ETByStoreKeys_sorted.findIndex(
                  (i) => i == 'Report Total'
                );
                if (index >= 0) {
                  ETByStoreKeys_sorted.splice(index, 1);
                  ETByStoreKeys_sorted.splice(0, 0, 'Report Total');
                }

                if (this.ShowHideBudget === 'Show') {
                  // Insert 'BDGT_TOTAL' and 'VARBDGT' after 'Report Total'
                  const reportTotalIndex = ETByStoreKeys_sorted.indexOf('Report Total');
                  if (reportTotalIndex >= 0) {  // Remove BDGT_TOTAL and VARBDGT if already in list
                    ETByStoreKeys_sorted = ETByStoreKeys_sorted.filter(
                      (x) => x !== 'BDGT_TOTAL' && x !== 'VARBDGT'
                    );

                    // Insert them in order
                    ETByStoreKeys_sorted.splice(reportTotalIndex + 1, 0, 'BDGT_TOTAL', 'VARBDGT');
                  }
                }
                const Cmtindex = ETByStoreKeys_sorted.findIndex(
                  (i) => i == 'CommentsStatus'
                );
                if (Cmtindex >= 0) {
                  ETByStoreKeys_sorted.splice(Cmtindex, 1);
                  // ETByStoreKeys_sorted.splice(0, 0, 'CommentsStatus');
                }
                for (var i = 0; i < ETByStoreKeys_sorted.length; i++) {
                  this.ExpenseTrendByStoreKeys.push(
                    ETByStoreKeys_sorted[i].toLowerCase()
                  );
                  this.AllDatakeys.push(ETByStoreKeys_sorted[i]);
                }
                this.ExpenseTrendByStoreKeys.push(AllStore_Label);
                this.AllDatakeys.push(AllStore_Label);
                let XpenseTrendByStoreData = x.response;
                let ETByStoreData_sorted = XpenseTrendByStoreData;
                let AllStoreLabel_Data = '';
                this.ExpenseTrendByStore = x.response;
                if (this.ExpenseTrendByStoreKeys.length == 0) {
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
              this.ExpenseTrendByStore = [];
              this.AllDatakeys = [];
              this.spinner.hide();
              this.NoData = true;
            }
          } else {
            let resTime = (new Date().getTime() - startFrom) / 1000;
            this.toast.show(
              x.status,
              'danger',
              'Error'
            );
            this.spinner.hide();
            this.NoData = true;
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
    this.apiSrvc.GetHeaderData().subscribe((res) => {
    });
  }
  getStores() {
    this.selectedstorevalues = [];
  }
  openDetails(Object: any, storename: any, ref: any, date: any) {
    console.log(Object, storename);
    if (ref == 'store') {
      const dateMonth =
        date.toString().substr(8, 2) +
        '-' +
        date.toString().substr(4, 3) +
        '-' +
        date.toString().substr(11, 4);
      const DetailsSF = this.ngbmodal.open(IncomestatementstoreDetails, {
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
        STORES: 2,
        LatestDate: dateMonth,
        Group: this.groups,
      };
    }
  }
  selBlock: any;
  commentopen(item: any, i: any, slblock: any = '') {
    this.index = '';
    //console.log('Selected Obj :', item);
    //return
    this.selBlock = slblock + i.toString();
    this.index = i.toString();
    this.commentobj = {
      TYPE: item.LABLE,
      NAME: item.LABLE,
      STORES: i,
      STORENAME: item.LABLE,
      Month: '',
      ModuleId: '73',
      ModuleRef: 'ISSC',
      state: 1,
      indexval: i,
    };
    // const DetailsSF = this.ngbmodal.open(CommentsComponent, {
    //   size: 'xl',
    //   backdrop: 'static',
    // });
    // DetailsSF.componentInstance.SFComments = {
    //   TYPE: item.LABLEVAL,
    //   NAME: item.LABLE,
    //   STORES: this.selectedstorevalues,
    //   LatestDate: this.Month,
    //   STORENAME: this.selectedstorename,
    //   Month: this.Month,
    //   ModuleId: '66',
    //   ModuleRef: 'SF',
    // };
    // DetailsSF.result.then(
    //   (data) => {},
    //   (reason) => {
    //     // alert('close');
    //     // // on dismiss
    //     // const Data = {
    //     //   state: true,
    //     // };
    //     // this.apiSrvc.setBackgroundstate({ obj: Data });
    //     this.GetData();
    //   }
    // );
    //alert(this.index);
  }
  index = '';
  commentobj: any = {};
  addcmt(data: any) {
    // //console.log('Checking Add cmt  : ', data);
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
            // //console.log(data);
          },
          (reason) => {
            // //console.log(reason);
            if (reason == 'O') {
              this.commentobj['state'] = 1;
              this.index = this.commentobj['indexval'];
            } else {
              this.commentobj['state'] = 1;
              this.index = this.commentobj['indexval'];
              if (this.Filter == 'StoreSummary') {
                this.GetData(this.currentMonth);
              }
            }
            // // on dismiss
            // const Data = {
            //   state: true,
            // };
            // this.apiSrvc.setBackgroundstate({ obj: Data });
          }
        );
    }
    if (data == 'AD') {
      if (this.Filter == 'StoreSummary') {
        this.GetData(this.currentMonth);
      }
    }
  }
  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    }
    return false;
  }
  SFstate: any;
  ngAfterViewInit(): void {
    this.apiSrvc.GetReportOpening().subscribe((res) => {
      // //console.log(res);
      // if (res.obj.Module == 'Income Statement Store') {
      //   const modalRef = this.ngbmodal.open(StoreReportComponent, {
      //     size: 'xl',
      //     backdrop: 'static',
      //   });
      // }
      if (res.obj.Module == 'Income Statement Store') {
        document.getElementById('report')?.click()
      }
    });
    this.apiSrvc.getExportToExcelAllReports().subscribe((res) => {
      this.SFstate = res.obj.state;
      if (res.obj.title == 'Income Statement Store') {
        if (res.obj.state == true) {
          this.exportAsXLSX();
        }
      }
    });
    this.apiSrvc.GetReports().subscribe((data) => {
      //console.log(data)
      if (data.obj.Reference == 'Income Statement Store') {
        if (data.obj.header == undefined) {
          this.date = data.obj.month;
          this.Month = data.obj.month;
          // this.Month =  ('0' +( new Date(data.obj.month).getMonth())).slice(-2)+'-01'+'-'+new Date (data.obj.month).getFullYear()
          this.StoreValues = data.obj.storeValues;
          this.StoreName = data.obj.Sname;
          this.Filter = data.obj.filters;
          this.SubFilter = data.obj.subfilters;
          this.ShowHideBudget = data.obj.ShowHideBudget

          this.groups = data.obj.groups;
          if (this.StoreValues.toString().indexOf(',') > 0) {
            this.gridvisibility = 'DL';
          } else {
            this.gridvisibility = 'SL';
          }
          // this.DataSelection(this.Filter);
        } else {
          if (data.obj.header == 'Yes') {
            this.StoreValues = data.obj.storeValues;
            //console.log(this.StoreValues);
          }
        }
        const headerdata = {
          title: 'Income Statement Store',
          path1: '',
          path2: '',
          path3: '',
          Month: this.Month,
          filter: this.Filter,
          stores: this.StoreValues,
          ShowHideBudget: this.ShowHideBudget,

          groups: this.groups,
        };
        this.apiSrvc.SetHeaderData({
          obj: headerdata,
        });
        // let date=    ('0' + ( this.Month.getMonth() + 1)).slice(-2) +
        // '-01' +
        // '-' +
        // this.Month.getFullYear();
        this.header = [{ type: 'Bar', storeIds: this.StoreValues, Month: this.Month, groups: this.groups, ShowHideBudget: this.ShowHideBudget, }]
        if (this.StoreValues != '') {
          this.GetData(this.currentMonth);
        } else {
          this.NoData = true;
          this.ExpenseTrendByStore = [];
          this.ExpenseTrendByStoreKeys = []
        }
      }
    });
    this.apiSrvc.getExportToPDFAllReports().subscribe((res: { obj: { state: boolean; title: string; }; }) => {
      if (res.obj.title == 'Income Statement Store') {
        if (res.obj.state == true) {
          this.generatePDF();
        }
      }
    });
    this.apiSrvc.getExportToPrintAllReports().subscribe((res: { obj: { state: boolean; title: string; }; }) => {
      if (res.obj.title == 'Income Statement Store') {
        if (res.obj.state == true) {
          this.GetPrintData();
        }
      }
    });
    this.apiSrvc.getExportToEmailPDFAllReports().subscribe((res: { obj: { state: boolean; title: string; }; }) => {
      if (res.obj.title == 'Income Statement Store') {
        if (res.obj.state == true) {
          // this.sendEmailData(res.obj.Email, res.obj.notes, res.obj.from);
        }
      }
    });
  }
  reportOpen(temp: any) {
    this.ngbmodalActive = this.ngbmodal.open(temp, {
      size: 'xl',
      backdrop: 'static',
    });
  }
  LogCount = 1;
  logSaving(url: any, object: any, time: any, status: any) {
    let ip = localStorage.getItem('Browser');
    // //console.log(object);
    const data = JSON.parse(localStorage.getItem('UserDetails')!);
    // //console.log(data);
    if (
      data != 'None' &&
      data != undefined &&
      data != null &&
      data != '' &&
      this.LogCount == 1
    ) {
      const obj = {
        UL_DealerId: '1',
        UL_GroupId: '',
        UL_UserId: data.userid,
        UL_IpAddress: ip!.split(',')[1],
        UL_Browser: ip!.split(',')[0],
        UL_Absolute_URL: window.location.href,
        UL_Api_URL: url,
        UL_Api_Request: JSON.stringify(object),
        UL_PageName: 'Income Statement Store',
        UL_ResponseTime: time,
        UL_Token: '',
        UL_ResponseStatus: status,
        UL_Groupings: '',
        UL_Timeframe: '',
        UL_Stores: '',
        UL_Filters: '',
        UL_Teams: '',
        UL_Status: 'Y',
      };
      this.apiSrvc.postmethod('useractivitylog', obj).subscribe((val) => {
        this.LogCount = 0;
      });
    }
  }
  close(data: any) {
    this.index = '';
  }


  isBoldRow(label: string): boolean {
    const boldLabels = [
      'Unit Retail Sales', 'Pure Gross', 'Variable Gross', 'Total Fixed',
      'Total Store Gross', 'Total Expenses', 'Operating Profit',
      'Net Adds/Deducts', 'Net Profit'
    ];
    return boldLabels.includes(label);
  }

  getCellClasses(item: any, itemKey: string) {
    const isNegative = !this.inTheGreen(item[itemKey]);
    const isBold = this.isBoldRow(item.LABLE);
    const isNonClickable = this.isNonClickableKey(itemKey);

    return {
      negative: isNegative && ['Operating Profit', 'Net Profit'].includes(item.LABLE),
      setPerens: isNegative,
      bold: isBold,
      notbold: !isBold,
      nopointerevents: isNonClickable,
      pointerevents: !isNonClickable
    };
  }

  isNonClickableKey(key: string): boolean {
    return key.startsWith('BDGT_') || key.startsWith('VAR') || key === 'Report Total';
  }


  exportAsXLSX(): void {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet(
      'Income Statement Store'
    );
    worksheet.views = [
      {
        state: 'frozen',
        ySplit: 13, // Number of rows to freeze (2 means the first two rows are frozen)
        topLeftCell: 'A14', // Specify the cell to start freezing from (in this case, the third row)
        showGridLines: false,
      },
    ];
    worksheet.addRow('');
    const titleRow = worksheet.addRow(['Income Statement Store']);
    titleRow.eachCell((cell, number) => {
      cell.alignment = { indent: 1, vertical: 'top', horizontal: 'left' };
    });
    titleRow.font = { name: 'Arial', family: 4, size: 12, bold: true };
    titleRow.worksheet.mergeCells('A2', 'D2');
    worksheet.addRow('');
    const DateToday = this.datepipe.transform(
      new Date(),
      'MM/dd/yyyy h:mm:ss a'
    );
    const DATE_EXTENSION = this.datepipe.transform(new Date(), 'MMddyyyy');
    worksheet.addRow([DateToday]).font = {
      name: 'Arial',
      family: 4,
      size: 9,
    };
    // ,ExpenseTrendByStore
    let ISstoresByData = this.ExpenseTrendByStore.map(
      (_arrayElement: any) => Object.assign({}, _arrayElement)
    );
    const MonthDate = this.datepipe.transform(this.Month, 'MMM-yy');

    // Create a copy and reorder keys: BDGT_TOTAL and VARBDGT should follow Report Total
    let storeHeadings = [...this.StoreNamesHeadings];

    storeHeadings = storeHeadings.filter((x) => {
      if (this.ShowHideBudget === 'Show') return true;

      // Hide all budget/variance keys
      const isDynamicBudget = /^BDGT_\d+$/.test(x);
      const isDynamicVariance = /^VAR\d+$/.test(x);
      const isBudgetTotal = x === 'BDGT_TOTAL';
      const isBudgetVariance = x === 'VARBDGT';

      return !isDynamicBudget && !isDynamicVariance && !isBudgetTotal && !isBudgetVariance;
    });

    if (this.ShowHideBudget === 'Show') {
      const reportTotalIndex = storeHeadings.indexOf('Report Total');
      if (reportTotalIndex >= 0) {
        // Ensure not duplicated
        storeHeadings = storeHeadings.filter(
          (k) => k !== 'BDGT_TOTAL' && k !== 'VARBDGT'
        );
        storeHeadings.splice(reportTotalIndex + 1, 0, 'BDGT_TOTAL', 'VARBDGT');
      }
    }
    const Header = [MonthDate];
    for (let i = 0; i < storeHeadings.length; i++) {
      Header.push(this.excelformatKey(storeHeadings[i]));
    }
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
    datemonth.value = this.Month;
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
    groups.value =
      'WesternAuto';
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
    stores.value = 'WesternAuto';
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
      cell.alignment = {
        vertical: 'top',
        horizontal: 'center',
        wrapText: true,
      };
    });
    ISstoresByData.forEach((d: any) => {
      var Obj = [d.LABLE];
      storeHeadings.forEach((e: any) => {
        Obj = [
          ...Obj,
          d[e] == '' ? '-' : d[e] == null ? '-' : parseInt(d[e]),
        ];
      });
      // //console.log(Obj);
      const row = worksheet.addRow(Obj);
      row.getCell(1).alignment = {
        indent: 1,
        vertical: 'middle',
        horizontal: 'left',
      };
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
    worksheet.getColumn(19).width = 15;
    worksheet.getColumn(20).width = 15;
    worksheet.getColumn(21).width = 15;
    worksheet.getColumn(22).width = 15;
    worksheet.getColumn(23).width = 15;
    worksheet.getColumn(24).width = 15;
    worksheet.getColumn(25).width = 15;
    worksheet.getColumn(26).width = 15;
    worksheet.getColumn(27).width = 25;
    worksheet.addRow([]);
    workbook.xlsx.writeBuffer().then((data: any) => {
      const blob = new Blob([data], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      FileSaver.saveAs(
        blob,
        'Income Statement Store Composite_' +
        DATE_EXTENSION +
        EXCEL_EXTENSION
      );
    });
    // });
  }

  generatePDF() {
    this.spinner.show();
    const printContents = document.getElementById('IncomeStatementStore')!.innerHTML;
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
      <title>Income Statement Store</title>
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
                     .performance-scorecard .table tr:nth-child(even) td:first-child,
                     .performance-scorecard .table tr:nth-child(even) td:nth-child(2) {
                      background-color: #ffffff;
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
        doc.save('Income Statement Store.pdf');
        // popupWin.close();
        this.spinner.hide();
      });
  }
  sendEmailData(Email: any, notes: any, from: any) {
    this.spinner.show();
    const printContents = document.getElementById('IncomeStatementStore')!.innerHTML;
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
              <title>Income Statement Store</title>
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
                     .performance-scorecard .table tr:nth-child(even) td:first-child,
                     .performance-scorecard .table tr:nth-child(even) td:nth-child(2) {
                      background-color: #ffffff;
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
      scale: 1 // Adjust scale to fit the page better
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
        const pdfFile = this.blobToFile(pdfBlob, 'Income Statement Store.pdf');
        const formData = new FormData();
        formData.append('to_email', Email);
        formData.append('subject', 'Income Statement Store');
        formData.append('file', pdfFile);
        formData.append('from', from);
        this.apiSrvc.postmethod(this.comm.routeEndpoint + 'mail', formData).subscribe(
          (res: any) => {
            console.log('Response:', res);
            if (res.status === 200) {
              // alert(res.response);
              this.toast.success(res.response)
            } else {
              alert('Invalid Details');
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