import { Component, ElementRef, HostListener, Injector, ViewChild } from '@angular/core';
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
import { BsDatepickerConfig, BsDatepickerDirective, BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { Router } from '@angular/router';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { Stores } from '../../../../CommonFilters/stores/stores';

@Component({
  selector: 'app-dashboard',
  imports: [SharedModule, BsDatepickerModule, Stores],
  standalone: true,
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss'
})
export class Dashboard {
  Current_Date: any;
  subscription!: Subscription;
  subscriptionExcel!: Subscription;
  subscriptionReport!: Subscription;
  subscriptionPDF!: Subscription;
  subscriptionEmailPDF!: Subscription;
  // gridshow:boolean=false
  // monthgridshow:boolean=false
  FSData: any = [];
  date: any;
  ExpenseTrendByStoreKeys: any = [];
  AllDatakeys: any = [];
  ExpenseTrendByStore_Excel: any;
  XpenseTrendByStoreKeys: any = [];
  ExpenseTrendByStore: any;
  ExpenseTrendByStore_ExcelMonth: any;
  XpenseTrendByStoreKeysMonth: any = [];
  ExpenseTrendByStoreKeysMonth: any = [];
  AllDatakeysMonth: any = [];
  ExpenseTrendByStoreMonth: any;
  SubFilter: any;
  SelectedTab: any = [];
  Month!: any;
  // stores: any;
  selectedstorevalues: any = [];
  selectedstorename: any;
  selectedFilters: string[] = [];
  selectedLabel: string = '( All )';
  StoreName: any = 'All Stores';
  Filter: any = ['New', 'Used', 'Service', 'Parts', 'Detail'];
  StoreValues: any = 2;
  groups: any = 1;
  PresentDayDate: string;

  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent) {
    const clickedInside = (event.target as HTMLElement).closest('.dropdown-toggle, .reportstores-card , .timeframe');
    if (!clickedInside) {
      this.activePopover = -1;
    }
  }


  stores: any = [];
  selectDate: Date = new Date();
  currentMonth: any = '';
  groupsArray: any = [];
  storename: any = ''
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  groupId: any = 0;
  storeIds: any = '0';
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
    private comm: common,
    private router: Router,
    private toast: ToastService,
    private injector: Injector,
    public shared: Sharedservice,
  ) {
    const lastMonth = new Date();
    let today = new Date();
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
    if (today.getDate() < 5) {
      this.date = new Date(lastMonth.setMonth(lastMonth.getMonth() - 1));
    } else {
      this.date = new Date(lastMonth.setMonth(lastMonth.getMonth()));
    }
    this.title.setTitle(this.comm.titleName + '-Enterprise Income / Expense Trend');
    if (localStorage.getItem('Fav') != 'Y') {
      const data = {
        title: 'Enterprise Income / Expense Trend',
        path1: '',
        path2: '',
        path3: '',
        Month: this.date,
        stores: this.storeIds.toString(),
        store: this.storeIds,
        filter: this.Filter,
        groups: 1,
        count: 0,
      };
      this.apiSrvc.SetHeaderData({
        obj: data,
      });
    }
    const format = 'ddMMyyyy';
    const locale = 'en-US';
    const myDate = new Date();
    const formattedDate = formatDate(myDate, format, locale);
    this.PresentDayDate = formattedDate;
    this.selectDate = this.date
    this.currentMonth = this.selectDate;
    this.selectedFilters = [...this.filters];
    this.GetDataByMonths(this.currentMonth, this.selectedFilters);
  }
  roleId: any;
  ngOnInit(): void {
    this.roleId = localStorage.getItem('roleId');
    console.log('role Id', this.roleId);
    this.getStores();

  }

  StoreNamesHeadings: any = [];
  MonthsHeadings: any = [];

  Scrollpercent: any = 0;
  scrollPositionStoring: number = 0;
  scrollCurrentPosition: number = 0;

  @ViewChild('scrollcent') scrollcent!: ElementRef;

  updateVerticalScroll(event: any): void {
    this.scrollCurrentPosition = event.target.scrollTop;
  }

  // Filters list
  filters: string[] = ['New', 'Used', 'Service', 'Parts', 'Detail'];
  activePopover: number | null = null;
  bsConfig: Partial<BsDatepickerConfig> = {
    dateInputFormat: 'MMMM/YYYY',
    minMode: 'month',
    maxDate: new Date()
  };
  monthPicker!: BsDatepickerDirective;
  openMonthPicker() {
    if (this.monthPicker) {
      this.monthPicker.show();
    }
  }

  togglePopover(index: number) {
    this.activePopover = this.activePopover === index ? null : index;
  }


  isSelected(filter: string): boolean {
    return this.selectedFilters.includes(filter);
  }

  toggleFilter(filter: string) {
    const index = this.selectedFilters.indexOf(filter);
    if (index >= 0) {
      this.selectedFilters.splice(index, 1);
    } else {
      this.selectedFilters.push(filter);
    }
    this.updateLabel();
  }

  updateLabel() {
    if (!this.selectedFilters || this.selectedFilters.length === 0) {
      this.selectedLabel = '( Select )';
    } else if (this.selectedFilters.length === this.filters.length) {
      this.selectedLabel = '( All )';
    } else if (this.selectedFilters.length === 1) {
      this.selectedLabel = `( ${this.selectedFilters[0]} )`;
    } else {
      this.selectedLabel = `( Selected ${this.selectedFilters.length} )`;
    }
  }



  applyDateChange() {
    if (!this.storeIds || this.storeIds.length === 0) {
      this.toast.show('Please Select Atleast One Store', 'warning', 'Warning');
      return;
    }
    if (!this.selectedFilters || this.selectedFilters.length === 0) {
      this.toast.show('Please Select Atleast One Department', 'warning', 'Warning');
      return;
    }
    else {
      this.currentMonth = this.formatMonth(this.selectDate);
      this.GetDataByMonths(this.currentMonth, this.selectedFilters);
      this.activePopover = null;
      this.isLoading = true;
    }
  }
  formatMonth(date: Date): string {
    const year = date.getFullYear();
    const month = ('0' + (date.getMonth() + 1)).slice(-2);
    return `${year}-${month}`;
  }
  isLoading = true;
  NoData = false;
  ShowHideGP: any = 'Hide';
  ShowHideBudget: any = 'Hide';
  EndDate: any;
  dates: any;
  storeName: any;
  ExpenseTrendKeys: any = [];
  IncomeSummaryDataKeys: any = [];
  ServiceTrendingSubKeys: any = [];
  ServiceTrendingXlm: any = [];
  IncomeSummaryData: any;
  GetDataByMonths(date: any, filters: any) {
    this.spinner.show();
    this.dates = [];
    this.IncomeSummaryData = [];
    const currentDate = new Date(date);
    const currentMonth = currentDate.getMonth();
    const currentYear = currentDate.getFullYear();
    for (let i = 0; i < 12; i++) {
      const lastMonthDate = new Date(currentYear, currentMonth - i, 1);
      this.dates.push(this.formatDate(lastMonthDate));
    }
    const DateToday = this.datepipe.transform(
      new Date(date),
      'yyyy-MM-dd'
    );
    const obj = {
      DATE: DateToday,
      AS_ID: this.storeIds.toString(),
      Variance: 'N'
    };
    console.log(obj);
    this.apiSrvc
      .postmethod(this.comm.routeEndpoint + 'GetDepartExpenseBudgetTrendV1', obj)
      .subscribe(
        (x: any) => {
          const currentTitle = document.title;
          if (x.status == 200 && x.response) {
            let data = x.response.reduce(
              (r: any, { FINSUMMARY, ...rest }: any) => {
                if (!r.some((o: any) => o.FINSUMMARY == FINSUMMARY)) {
                  r.push({
                    FINSUMMARY,
                    ...rest,
                    subdata: x.response.filter(
                      (v: any) => v.FINSUMMARY == FINSUMMARY
                    ),
                  });
                }
                return r;
              },
              []
            );
            this.IncomeSummaryData = data;
            console.log('IncomeSummaryData', this.IncomeSummaryData);

            const IncomeSummaryKeys = Object.keys(x.response[0] || {}).slice(6);
            this.IncomeSummaryDataKeys = IncomeSummaryKeys;
            this.IncomeSummaryDataKeys = this.IncomeSummaryDataKeys.filter((dealership: any) =>
              dealership !== 'TOTAL' &&
              dealership !== 'AVG' &&
              dealership !== 'SEQ' &&
              (this.ShowHideGP === 'Show' || (this.ShowHideGP === 'Hide' && !dealership.endsWith('_GP'))) &&
              (this.ShowHideBudget === 'Show' ||
                (this.ShowHideBudget === 'Hide' &&
                  !dealership.endsWith('_BDGT') &&
                  !dealership.startsWith('VAR')))
            );



            if (this.IncomeSummaryData.length > 0) {
              this.IncomeSummaryData.forEach((x: any) => {
                (x.SubData = []), (x.data2sign = '+');
              });

              console.log('Keys ', this.IncomeSummaryDataKeys);
              this.NoData = this.IncomeSummaryDataKeys.length === 0;
            } else {
              this.NoData = true;
            }
          } else {
            this.NoData = true;
          }

          this.spinner.hide();
        },
        (error) => {
          console.error(error);
          this.NoData = true;
          this.spinner.hide();
        }
      );
  }
  formatDate(date: Date): string {
    const options: Intl.DateTimeFormatOptions = {
      year: 'numeric',
      month: 'short',
    };
    return date.toLocaleDateString(undefined, options);
  }
  formatKey(key: string): string {
    if (key.endsWith('_BDGT')) return 'BUDGET';
    if (key.startsWith('Varience')) return 'VARIANCE';
    return key;
  }
  isKeyVisible(key: string): boolean {
    return this.IncomeSummaryDataKeys.includes(key);
  }
  isKeyVisibles(key: string): boolean {
    return key !== '-1';

  }
  disabledSubTypes = [
    'Variable Gross',
    'Fixed Gross',
    'New Incentive Gross',
    'Total Store Gross',
    'Net Profit'
  ];
  isDisabled(item: any, key: string): boolean {
    return (
      item[key] == 0 ||
      item[key] == null ||
      item[key] === '' ||
      this.disabledSubTypes.includes(item['SUBTYPE DETAIL'])
    );
  }
  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    }
    return false;
  }

  SFstate: any;
  StoreCodes: any;
  block: any = '';
  Report: any = '';
  ngAfterViewInit(): void {
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.shared.common.pageName == 'Enterprise Income / Expense Trend') {
        if (res.obj.storesData != undefined) {
          this.groupsArray = res.obj.storesData;
          this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
          this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_name : this.groupName = ''
          this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
          this.getStoresandGroupsValues()
        }
      }
    })
    this.subscriptionReport = this.apiSrvc
      .GetReportOpening()
      .subscribe((res) => {
        console.log(res);
        if (this.subscriptionReport != undefined) {
          if (res.obj.Module == 'Enterprise Income / Expense Trend') {
            document.getElementById('report')?.click();
          }
        }
      });
    this.apiSrvc.GetReports().subscribe((data) => {
      if (data.obj.Reference == 'Enterprise Income / Expense Trend') {
        console.log(data);
        if (data.obj.header == undefined) {
          this.date = data.obj.month;
          this.Month = data.obj.month;
          this.StoreValues = data.obj.storeValues;
          this.StoreName = data.obj.Sname;
          this.Filter = data.obj.filter;
          this.SubFilter = data.obj.subfilters;
          this.StoreCodes = data.obj.storecode;
          this.groups = data.obj.groups;
          this.index = '';
          this.Scrollpercent = 0;
        } else {
          if (data.obj.header == 'Yes') {
            this.StoreValues = data.obj.storeValues;
          }
        }
        this.GetDataByMonths(this.currentMonth, this.selectedFilters);
        if (this.StoreValues != '') {
          this.goToFirstPage();
        } else {
          this.NoData = true;
          this.filteredETdetailsData = [];
        }
        const headerdata = {
          title: 'Enterprise Income / Expense Trend',
          path1: '',
          path2: '',
          path3: '',
          Month: new Date(this.Month),
          filter: this.Filter,
          stores: this.StoreValues,
          storecode: this.StoreCodes,
          Sname: this.StoreName,
          groups: this.groups,
        };
        this.apiSrvc.SetHeaderData({
          obj: headerdata,
        });
      }
    });
    this.subscriptionExcel = this.apiSrvc
      .getExportToExcelAllReports()
      .subscribe((res) => {
        this.SFstate = res.obj.state;
        if (this.subscriptionExcel != undefined) {
          if (res.obj.title == 'Enterprise Income / Expense Trend') {
            if (res.obj.state == true) {
              if (this.ExpenseTrend == true) {
                this.exportAsXLSX();
              } else if (this.ExpenseTrendDetails == true) {
                this.exportToExcel();
              }
            }
          }
        }
      });
    this.subscription = this.apiSrvc
      .getExportToPrintAllReports()
      .subscribe((res) => {
        if (this.subscription != undefined) {
          if (res.obj.title == 'Enterprise Income / Expense Trend') {
            if (res.obj.statePrint == true) {
              this.GetPrintData();
            }
          }
        }
      });
    this.subscriptionPDF = this.apiSrvc
      .getExportToPDFAllReports()
      .subscribe((res) => {
        if (this.subscriptionPDF != undefined) {
          if (res.obj.title == 'Enterprise Income / Expense Trend') {
            if (res.obj.statePDF == true) {
              if (this.ExpenseTrend == true) {
                this.generatePDF();
              } else if (this.ExpenseTrendDetails == true) {
                this.generatePDFDetails();
              }
            }
          }
        }
      });
    this.subscriptionEmailPDF = this.apiSrvc
      .getExportToEmailPDFAllReports()
      .subscribe((res) => {
        if (this.subscriptionEmailPDF != undefined) {
          if (res.obj.title == 'Enterprise Income / Expense Trend') {
            if (res.obj.stateEmailPdf == true) {
              if (this.ExpenseTrend == true) {
                this.sendEmailData(res.obj.Email, res.obj.notes, res.obj.from);
              } else if (this.ExpenseTrendDetails == true) {
                this.sendEmailDataDetails(
                  res.obj.Email,
                  res.obj.notes,
                  res.obj.from
                );
              }
            }
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
  closeReport() {
    this.Report = '';
  }
  getStores() {
    this.selectedstorevalues = [];
    this.stores = JSON.parse(localStorage.getItem('Stores')!);
  }
  ETdetailsData: any = [];
  currentPage: number = 1;
  itemsPerPage: number = 500;
  maxPageButtonsToShow: number = 3;
  clickedPage: number | null = null;
  filteredETdetailsData: any[] = [];
  selectedDate: any;
  DateType: any;
  DetailsSearchName: any;
  Lable: any;
  searchText: string = '';
  ExpenseTrend: boolean = true;
  ExpenseTrendDetails: boolean = false;
  SubtypeDetailLable: any;
  FinSummaryLable: any;
  Dept: any;
  MonthDate: any;
  openMonthDetails(
    Object: any,
    date: any,
    ref: any,
    item: any,
    DateMethod: any
  ) {
    console.log('Object', Object);
    console.log('date', date);
    console.log('ref', ref);
    console.log('item', item);
    console.log('DateMethod', DateMethod);
    this.spinner.show();
    this.scrollPositionStoring = this.scrollCurrentPosition;
    this.ExpenseTrend = false;
    this.ExpenseTrendDetails = true;
    // this.ETdetailsData = [];
    this.DateType = DateMethod;
    this.Lable = ref;
    console.log(date);
    this.MonthDate = date;
    if (DateMethod === 'YTD') {
      this.selectedDate = this.datepipe.transform(new Date(date), 'yyyy');
    } else if (DateMethod === 'PYTD') {
      const currentDate = new Date(date);
      const previousYearDate = new Date(
        currentDate.setFullYear(currentDate.getFullYear() - 1)
      );
      this.selectedDate = this.datepipe.transform(previousYearDate, 'yyyy');
    } else {
      this.selectedDate = this.datepipe.transform(new Date(date), 'MM-yyyy');
    }
    this.SubtypeDetailLable = Object['SUBTYPE DETAIL'];
    this.FinSummaryLable = Object.DISPLAY_PARENTLABLE;
    this.Dept = Object.Category
    const Obj = {
      as_ids: this.storeIds.toString(),
      // dept: this.Filter.toString(),
      dept: Object.Category,
      subtype: '',
      subtypedetail: Object['SUBTYPE DETAIL'],
      FinSummary: '',
      date: this.selectedDate,
      accountnumber: ''
    };
    console.log(Obj);
    this.apiSrvc
      .postmethod(this.comm.routeEndpoint + 'GetExpenseTrendDetailsV1', Obj)
      .subscribe((res) => {
        if (res.status == 200) {
          this.ETdetailsData = res.response;
          this.filterData();
          console.log('ET Details', this.ETdetailsData);
          this.spinner.hide();
          this.NoData = true;
        }
      });
  }
  expandedIndex: number | null = null;
  FSSubDetailsMap: { [index: number]: any[] } = {};

  GetSubDetails(AcctNo: any, StoreName: any, index: number) {
    if (this.expandedIndex === index) {
      this.expandedIndex = null;
      return;
    }
    this.expandedIndex = index;
    this.spinner.show();
    const Obj = {
      as_ids: StoreName,
      // dept: this.Filter.toString(),
      dept: this.Dept,
      subtype: '',
      subtypedetail: this.SubtypeDetailLable,
      FinSummary: '',
      date: this.selectedDate,
      accountnumber: AcctNo
    };

    this.apiSrvc
      .postmethod(this.comm.routeEndpoint + 'GetExpenseTrendDetailsV1', Obj)
      .subscribe((res) => {
        this.spinner.hide();
        if (res.status === 200) {
          this.FSSubDetailsMap[index] = res.response;
        }
      });
  }

  backtoWR() {
    this.ExpenseTrend = true;
    this.ExpenseTrendDetails = false;
    this.filteredETdetailsData = [];
    this.ETdetailsData = [];
    this.expandedIndex = null;
    if (this.StoreValues != '') {
      this.goToFirstPage();
    }
    setTimeout(() => {
      if (this.scrollcent && this.scrollcent.nativeElement) {
        this.scrollcent.nativeElement.scrollTop = this.scrollPositionStoring;
      }
    });
  }

  get postingAmountTotal(): number {
    return this.filteredETdetailsData.reduce((total, item) => {
      return total + (item.PostingAmount || 0);
    }, 0);
  }
  getPostingSubAmountTotal(index: number): number {
    const subRows = this.FSSubDetailsMap[index];
    if (!subRows || !Array.isArray(subRows)) {
      return 0;
    }

    return subRows.reduce((total, item) => {
      return total + (item.PostingAmount || 0);
    }, 0);
  }
  filterData() {
    const text = this.searchText.trim().toLowerCase();

    if (!text) {
      this.filteredETdetailsData = [...this.ETdetailsData];
    } else {
      this.filteredETdetailsData = this.ETdetailsData.filter((item: any) =>
        [
          item.StoreName,
          item.AccountNumber,
          item.AccountDescription,
          item.PostingAmount
        ].some(val =>
          val?.toString().toLowerCase().includes(text)
        )
      );
    }

    this.currentPage = 1;
  }
  get paginatedItems() {
    const start = (this.currentPage - 1) * this.itemsPerPage;
    return this.filteredETdetailsData.slice(start, start + this.itemsPerPage);
  }

  getMaxPageNumber(): number {
    return Math.max(1,
      Math.ceil(this.filteredETdetailsData.length / this.itemsPerPage)
    );
  }

  nextPage() {
    if (this.currentPage < this.getMaxPageNumber()) this.currentPage++;
  }

  prevPage() {
    if (this.currentPage > 1) this.currentPage--;
  }

  goToFirstPage() {
    this.currentPage = 1;
  }

  goToLastPage() {
    this.currentPage = this.getMaxPageNumber();
  }

  getStartRecordIndex(): number {
    return (this.currentPage - 1) * this.itemsPerPage + 1;
  }

  getEndRecordIndex(): number {
    return Math.min(
      this.getStartRecordIndex() + this.itemsPerPage - 1,
      this.filteredETdetailsData.length
    );
  }
  onPageSizeChange() {
    this.currentPage = 1;
  }

  openStoresDetails(Object: any, StoreName: any, date: any, item: any) {
    this.index = '';
    console.log(Object);
    console.log(item);
    const myDate = this.datepipe.transform(new Date(date), 'MMM-yyyy');
    let index = this.stores.filter(
      (store: any) => store.DEALER_NAME == StoreName
    );
    this.selectedstorevalues = index[0].AS_ID;
    this.selectedstorename = index[0].DEALER_NAME;
    const DetailsSF = this.ngbmodal.open({
      size: 'xl',
      backdrop: 'static',
    });
    DetailsSF.componentInstance.ETdetails = {
      TYPE: item.LABLEVAL,
      NAME: item.LABLE,
      STORES: this.selectedstorevalues,
      LatestDate: myDate,
      STORENAME: this.selectedstorename,
    };
  }
  openStoreGraph(
    Object: any,
    StoreName: any,
    date: any,
    item: any,
    SummaryType: any
  ) {
    console.log(Object, StoreName, date, item, SummaryType);
    const DetailsSF = this.ngbmodal.open({
      size: 'xl',
      backdrop: 'static',
    });
    var DisplayName = item.DISPLAY_LABLE;
    DetailsSF.componentInstance.ETgraphdetails = {
      ITEM: item,
      TYPE: item.DISPLAY_LABLE,
      NAME: SummaryType,
      SUMMARYTYPE: SummaryType,
    };
  }

  index = '';
  commentobj: any = {};
  close(data: any) {
    console.log(data);
    this.index = '';
  }
  addcmt(data: any) {
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
            console.log(data);
          },
          (reason) => {
            console.log(reason);

            if (reason == 'O') {
              this.commentobj['state'] = 1;
              this.index = this.commentobj['indexval'];
            } else {
              this.commentobj['state'] = 1;
              this.index = this.commentobj['indexval'];
              if (this.Filter == 'VariableTrendsvsBudget') {
                this.GetDataByMonths(this.currentMonth, this.selectedFilters);
              }
            }
            // // on dismiss

            // const Data = {
            //   state: true,
            // };
            // this.apiSrvc.setBackgroundstate({ obj: Data });
            // this.GetData();
          }
        );
    }
    if (data == 'AD') {
      if (this.Filter == 'VariableTrendsvsBudget') {
        this.GetDataByMonths(this.currentMonth, this.selectedFilters);
      }
    }
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
      title: 'Enterprise Income / Expense Trend',
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
    let storeNames: any = [];
    // let store = this.StoreValues.split(',');
    // storeNames = this.comm.groupsandstores
    //   .filter((v: any) => v.sg_id == this.groups)[0]
    //   .Stores.filter((item: any) =>
    //     store.some((cat: any) => cat === item.ID.toString())
    //   );
    // console.log(storeNames);

    // if (
    //   store.length ==
    //   this.comm.groupsandstores.filter((v: any) => v.sg_id == this.groups)[0]
    //     .Stores.length
    // ) {
    //   this.ExcelStoreNames = 'All Stores';
    // } else {
    //   this.ExcelStoreNames = storeNames.map(function (a: any) {
    //     return a.storename;
    //   });
    // }
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet('Enterprise Income By Expense Trend');
    worksheet.views = [
      {
        state: 'frozen',
        ySplit: 13,
        topLeftCell: 'A14',
        showGridLines: false,
      },
    ];
    worksheet.addRow('');
    const titleRow = worksheet.addRow(['Enterprise Income / Expense Trend']);
    titleRow.eachCell((cell, number) => {
      cell.alignment = { indent: 1, vertical: 'middle', horizontal: 'left' };
    });
    titleRow.font = { name: 'Arial', family: 4, size: 12, bold: true };
    titleRow.worksheet.mergeCells('A2', 'D2');

    worksheet.addRow('');
    const DateToday = this.datepipe.transform(
      new Date(),
      'MM.dd.yyyy h:mm:ss a'
    );
    worksheet.addRow([DateToday]).font = { name: 'Arial', family: 4, size: 9 };
    const ExpenseTrendByStoreMonth = this.ExpenseTrendByStoreMonth.map(
      (_arrayElement: any) => Object.assign({}, _arrayElement)
    );
    const StartDate = this.datepipe.transform(this.dates[0], 'MMM yyyy');
    const EndDate = this.datepipe.transform(this.dates[11], 'MMM yyyy');
    const Header = [StartDate + ' - ' + EndDate, 'YTD', 'PYTD', 'PACE', 'AVG'];
    for (let i = 0; i < this.ExpenseTrendKeys.length; i++) {
      Header.push(this.ExpenseTrendKeys[i]);
    }
    const ReportFilter = worksheet.addRow(['Report Filters :']);
    ReportFilter.font = { name: 'Arial', family: 4, size: 10, bold: true };

    const SummaryType = worksheet.addRow(['Summary Type :']);
    const summarytype = worksheet.getCell('B6');
    summarytype.value = 'Month Summary';
    summarytype.font = { name: 'Arial', family: 4, size: 9 };
    summarytype.alignment = { vertical: 'middle', horizontal: 'left' };
    SummaryType.getCell(1).font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
    };
    const DateMonth = worksheet.addRow(['Time Frame :']);
    const datemonth = worksheet.getCell('B7');
    datemonth.value = StartDate;
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
    // const groups = worksheet.getCell('B8');
    // groups.value = this.comm.groupsandstores.filter(
    //   (val: any) => val.sg_id == this.groups.toString()
    // )[0].sg_name;
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
    Stores.value = 'Stores :';
    const stores = worksheet.getCell('B9');
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
    worksheet.addRow('');
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
        fgColor: { argb: '2a91f0' },
        bgColor: { argb: 'FF0000FF' },
      };
      cell.border = { right: { style: 'thin' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });
    ExpenseTrendByStoreMonth.forEach((d: any) => {
      let Obj = [
        d.DISPLAY_LABLE,
        d.YTD == '' ? '-' : d.YTD == null ? '-' : d.YTD,
        d.PYTD == '' ? '-' : d.PYTD == null ? '-' : d.PYTD,
        d.PACE == '' ? '-' : d.PACE == null ? '-' : d.PACE,
        d.AVG == '' ? '-' : d.AVG == null ? '-' : d.AVG,
      ];
      this.ExpenseTrendKeys.forEach((e: any) => {
        Obj.push(d[e] === '' || d[e] === null ? '-' : parseInt(d[e], 10));
      });

      const row = worksheet.addRow(Obj);

      // Cell alignment for the first column
      row.getCell(1).alignment = {
        indent: 1,
        vertical: 'middle',
        horizontal: 'left',
      };

      if (row.number % 2 === 0) {
        row.eachCell((cell) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'e5e5e5' },
            bgColor: { argb: 'FF0000FF' },
          };
        });
      }
      if (d.DISPLAYHEAD_FLAG === 1) {
        row.eachCell((cell, colNumber) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: '94b6d1' },
          };
        });
        row.eachCell((cell, colNumber) => {
          if (colNumber === 1) {
            cell.font = {
              name: 'Arial',
              family: 4,
              size: 9,
              bold: true,
              color: { argb: 'FFFFFF' },
            };
          } else {
            cell.value = null;
          }
          cell.font = {
            name: 'Arial',
            family: 4,
            size: 9,
            bold: true,
            color: { argb: 'FFFFFF' },
          };
        });
      }
      if (d.ISHEAD_TOTAL === 'Y') {
        row.eachCell((cell) => {
          cell.alignment = {
            indent: 2,
            vertical: 'top',
            horizontal: 'right',
          };
          cell.font = {
            name: 'Arial',
            family: 4,
            size: 8,
            bold: true,
          };
        });
      }
      if (d.ISHEAD_TOTAL === 'N') {
        row.eachCell((cell) => {
          cell.alignment = {
            indent: 3,
            vertical: 'top',
            horizontal: 'right',
          };
        });
      } else {
        // Default alignment and font properties for all cells
        row.eachCell((cell) => {
          cell.border = { right: { style: 'thin' } };
          cell.alignment = {
            vertical: 'top',
            horizontal: 'right',
            indent: 1,
          };
        });
      }

      // Font settings for all cells
      row.font = { name: 'Arial', family: 4, size: 8 };

      // Border and alignment settings for all cells
      row.eachCell((cell) => {
        cell.border = { right: { style: 'thin' } };
        cell.alignment = {
          vertical: 'top',
          horizontal: 'right',
          indent: 1,
        };
      });

      // Conditional number formatting based on DISPLAY_LABEL
      if (d.DISPLAY_LABLE && d.DISPLAY_LABLE.includes('%')) {
        row.eachCell({ includeEmpty: true }, (cell, rowNumber) => {
          if (rowNumber > 1) {
            cell.numFmt = '0%';
          }
        });
      } else {
        row.eachCell({ includeEmpty: true }, (cell, rowNumber) => {
          if (rowNumber > 1) {
            cell.numFmt = '$#,##0';
          }
        });
      }
    });
    worksheet.getColumn(1).width = 30;
    worksheet.getColumn(1).alignment = {
      indent: 1,
      vertical: 'middle',
      horizontal: 'left',
    };

    for (let i = 2; i <= 28; i++) {
      worksheet.getColumn(i).width = 15;
    }
    worksheet.addRow([]);
    workbook.xlsx.writeBuffer().then((data: any) => {
      const blob = new Blob([data], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      FileSaver.saveAs(blob, 'Enterprise Income / Expense Trend' + EXCEL_EXTENSION);
    });
  }

  Favreports: any = [];

  isDesc: boolean = false;
  column: string = 'CategoryName';

  sort(property: string, data: any[], state?: any) {
    if (state === undefined) {
      this.isDesc = this.column === property ? !this.isDesc : false;
    }
    this.column = property;
    let direction = this.isDesc ? 1 : -1;
    data.sort((a, b) => {
      if (a[property] < b[property]) {
        return -1 * direction;
      } else if (a[property] > b[property]) {
        return 1 * direction;
      } else {
        return 0;
      }
    });
  }


  GetPrintData() {
    window.print();
  }
  generatePDF() {
    this.spinner.show();
    const printContents = document.getElementById('ExpenseTrend')!.innerHTML;
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
    <title>Enterprise Income / Expense Trend</title>
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
   .bg-white {
     background: #ffffff !important;
 }
   .negative {
     color: red;
 }
   .justify-content-between {
     justify-content: space-between !important;
 }
 .d-flex {
     display: flex !important;
 }
   .performance-scorecard  .table>:not(:first-child){
     border-top:0px solid #ffa51a
   }
   .performance-scorecard .table{
     text-align: end;
     text-transform: capitalize;
     border: transparent;

     width: 100%;
    } 
   .performance-scorecard .table .table th,
   .performance-scorecard .table td{
     white-space: nowrap;
     vertical-align: top;
   }
   .performance-scorecard .table th:first-child,
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
     text-align: center;
     font-family: 'FaktPro-Bold';
     font-size: 0.8rem;
    }
    .performance-scorecard .table thead th {
     padding: 5px 8px;
     margin: 0px; 
    }
    .performance-scorecard .table thead .bdr-btm{ 
     border-bottom: #005fa3;
    }
    .performance-scorecard .table thead tr:nth-child(1) { background-color: #fff !important;color: #000;text-transform: uppercase; border-bottom: #cfd6de;}
    .performance-scorecard .table thead tr:nth-child(2) { background-color: #337ab7 !important;color: #fff;text-transform: uppercase; border-bottom: #cfd6de;box-shadow: inset 0 1px 0 0 #cfd6de;} 
    .performance-scorecard .tablethead tr:nth-child(3) { background-color: #337ab7 !important;color: #fff;text-transform: uppercase; border-bottom: #cfd6de;box-shadow: inset 0 1px 0 0 #cfd6de;}
    
    .performance-scorecard .table tbody{
     font-family: 'FaktPro-Normal';
     font-size: .9rem;
    }
    .performance-scorecard .table tbody td{padding:2px 8px;margin: 0px; border: 1px solid #cfd6de }
    .performance-scorecard .table tbody tr{border-bottom: 1px solid #37a6f8;border-left: 1px solid #37a6f8}
    .performance-scorecard .table tbody td:first-child{text-align: start;box-shadow:inset -1px 0 0 0 #cfd6de ;}

    .performance-scorecard .table tbody .text-bold{font-family: 'FaktPro-Bold';}
    .performance-scorecard .table tbody .darkred-bg{ background-color: #282828 !important;color: #fff; }
    .performance-scorecard .table tbody .lightblue-bg{
      background-color: #94b6d1 !important;
      color: #000 !important;
      font-family: "FaktPro-Bold";
      padding-left: 1rem !important;
      pointer-events: none;}
    .performance-scorecard .table tbody .gold-bg{ background-color: #ffa51a;color: #fff;}
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
        doc.save('Enterprise Income / Expense Trend.pdf');
        // popupWin.close();
        this.spinner.hide();
      });
  }

  sendEmailData(Email: any, notes: any, from: any) {
    this.spinner.show();
    const printContents = document.getElementById('ExpenseTrend')!.innerHTML;
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
            <title>Enterprise Income / Expense Trend</title>
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
           .bg-white {
             background: #ffffff !important;
         }
           .negative {
             color: red;
         }
           .justify-content-between {
             justify-content: space-between !important;
         }
         .d-flex {
             display: flex !important;
         }
           .performance-scorecard  .table>:not(:first-child){
             border-top:0px solid #ffa51a
           }
           .performance-scorecard .table{
             text-align: end;
             text-transform: capitalize;
             border: transparent;
        
             width: 100%;
            } 
           .performance-scorecard .table .table th,
           .performance-scorecard .table td{
             white-space: nowrap;
             vertical-align: top;
           }
           .performance-scorecard .table th:first-child,
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
             text-align: center;
             font-family: 'FaktPro-Bold';
             font-size: 0.8rem;
            }
            .performance-scorecard .table thead th {
             padding: 5px 8px;
             margin: 0px; 
            }
            .performance-scorecard .table thead .bdr-btm{ 
             border-bottom: #005fa3;
            }
            .performance-scorecard .table thead tr:nth-child(1) { background-color: #fff !important;color: #000;text-transform: uppercase; border-bottom: #cfd6de;}
            .performance-scorecard .table thead tr:nth-child(2) { background-color: #337ab7 !important;color: #fff;text-transform: uppercase; border-bottom: #cfd6de;box-shadow: inset 0 1px 0 0 #cfd6de;} 
            .performance-scorecard .tablethead tr:nth-child(3) { background-color: #337ab7 !important;color: #fff;text-transform: uppercase; border-bottom: #cfd6de;box-shadow: inset 0 1px 0 0 #cfd6de;}
            
            .performance-scorecard .table tbody{
             font-family: 'FaktPro-Normal';
             font-size: .9rem;
            }
            .performance-scorecard .table tbody td{padding:2px 8px;margin: 0px; border: 1px solid #cfd6de }
            .performance-scorecard .table tbody tr{border-bottom: 1px solid #37a6f8;border-left: 1px solid #37a6f8}
            .performance-scorecard .table tbody td:first-child{text-align: start;box-shadow:inset -1px 0 0 0 #cfd6de ;}
        
            .performance-scorecard .table tbody .text-bold{font-family: 'FaktPro-Bold';}
            .performance-scorecard .table tbody .darkred-bg{ background-color: #282828 !important;color: #fff; }
            .performance-scorecard .table tbody .lightblue-bg{
              background-color: #94b6d1 !important;
              color: #000 !important;
              font-family: "FaktPro-Bold";
              padding-left: 1rem !important;
              pointer-events: none;}
            .performance-scorecard .table tbody .gold-bg{ background-color: #ffa51a;color: #fff;}
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
        const pdfFile = this.blobToFile(pdfBlob, 'Enterprise Income / Expense Trend.pdf');
        const formData = new FormData();
        formData.append('to_email', Email);
        formData.append('subject', 'Enterprise Income / Expense Trend');
        formData.append('file', pdfFile);
        formData.append('notes', notes);
        formData.append('from', from);
        this.apiSrvc
          .postmethod(this.comm.routeEndpoint + 'mail', formData)
          .subscribe(
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

  exportToExcel() {
    const FSDetailsData = [...this.filteredETdetailsData];
    const FSSubDetailsMap = this.FSSubDetailsMap;

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

    // Setup Excel
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet('Enterprise Income By Expense Trend Details');
    const DATE_EXTENSION = this.datepipe.transform(new Date(), 'MMddyyyy');
    const DateToday = this.datepipe.transform(new Date(), 'MM.dd.yyyy h:mm:ss a');

    worksheet.views = [{ state: 'frozen', ySplit: 10, topLeftCell: 'A11', showGridLines: false }];

    // Header section (above grid)
    worksheet.addRow([]);
    const titleRow = worksheet.addRow(['Enterprise Income / Expense Trend Details']);
    titleRow.font = { bold: true, size: 12 };
    worksheet.addRow([]);
    worksheet.addRow([DateToday]).font = { size: 9 };
    worksheet.addRow(['Selected Details:']).font = { bold: true, size: 10 };
    worksheet.addRow(['Type:', this.Lable]);
    worksheet.addRow(['Date:', this.selectedDate]);
    worksheet.addRow(['Store:', storeValue]);
    worksheet.addRow([]);

    // Grid Header
    const headers = [
      'Store Name',
      'Account Number',
      'Description',
      `Balance`,
    ];
    const headerRow = worksheet.addRow(headers);
    headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 9 };
    headerRow.eachCell((cell) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '2a91f0' },
      };
      cell.alignment = { horizontal: 'center', vertical: 'middle', indent: 1 };
    });

    let totalBalance = 0;
    const columnWidths = new Array(headers.length).fill(10);

    // Main Data Rows
    FSDetailsData.forEach((item: any, index: number) => {
      const rowData = [
        item.StoreName || '-',
        item.AccountNumber || '-',
        item.AccountDescription || '-',
        item.PostingAmount ?? '-',
      ];
      const mainRow = worksheet.addRow(rowData);
      mainRow.font = { size: 9 };

      if (typeof item.PostingAmount === 'number') {
        mainRow.getCell(4).numFmt = '$#,##0.00';
        totalBalance += item.PostingAmount;
      }
      mainRow.getCell(4).alignment = { horizontal: 'right', indent: 1 };

      mainRow.eachCell((cell) => {
        cell.alignment = { ...cell.alignment, indent: 1 };
      });

      if (index % 2 !== 0) {
        mainRow.eachCell((cell) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'e5e5e5' },
          };
        });
      }

      rowData.forEach((val, i) => {
        const length = val?.toString().length || 0;
        columnWidths[i] = Math.max(columnWidths[i], length);
      });

      const subDetails = FSSubDetailsMap[index];

      if (subDetails?.length && this.expandedIndex != null) {
        // Add Subtable Header
        const Subheaders = [
          'Control',
          'Date',
          'Detail Description',
          `Balance`,
        ];
        const subheaderRow = worksheet.addRow(Subheaders);
        subheaderRow.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 9 };
        subheaderRow.eachCell((cell) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: '2a91f0' },
          };
          cell.alignment = { horizontal: 'center', vertical: 'middle', indent: 1 };
        });

        // Add Sub Rows
        let subTotal = 0;

        subDetails.forEach((sub) => {
          const subData = [
            sub.Control || '-',
            sub.AccountingDate || '-',
            sub.DetailDescription || '-',
            sub.PostingAmount ?? '-',
          ];
          const subRow = worksheet.addRow(subData);
          subRow.font = { size: 9 };

          if (typeof sub.PostingAmount === 'number') {
            subTotal += sub.PostingAmount;
            totalBalance += sub.PostingAmount;
            subRow.getCell(4).numFmt = '$#,##0.00';
          }
          subRow.getCell(4).alignment = { horizontal: 'right', indent: 1 };

          subRow.eachCell((cell) => {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'd2deed' },
            };
            cell.alignment = { ...cell.alignment, indent: 1 };
          });

          subData.forEach((val, i) => {
            const length = val?.toString().length || 0;
            columnWidths[i] = Math.max(columnWidths[i], length);
          });
        });

        // Add Subtotal Row
        const subTotalRow = worksheet.addRow(['', '', 'Sub Total:', subTotal]);
        subTotalRow.font = { bold: true };
        subTotalRow.getCell(4).numFmt = '$#,##0.00';
        subTotalRow.getCell(4).alignment = { horizontal: 'right', indent: 1 };
        subTotalRow.eachCell((cell) => {
          cell.alignment = { ...cell.alignment, indent: 1 };
        });
      }
    });

    // Final Total Row
    const totalRow = worksheet.addRow(['', '', 'Total:', this.postingAmountTotal]);
    totalRow.font = { bold: true };
    totalRow.getCell(4).numFmt = '$#,##0.00';
    totalRow.getCell(4).alignment = { horizontal: 'right', indent: 1 };
    totalRow.eachCell((cell) => {
      cell.alignment = { ...cell.alignment, indent: 1 };
    });

    // Set Column Widths
    columnWidths.forEach((width, i) => {
      worksheet.getColumn(i + 1).width = Math.max(15, width + 2);
    });

    // Export file
    workbook.xlsx.writeBuffer().then((data) => {
      const blob = new Blob([data], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      FileSaver.saveAs(blob, `Expense_Trend_Details_${DATE_EXTENSION}.xlsx`);
    });
  }
  generatePDFDetails() {
    this.spinner.show();
    const printContents = document.getElementById(
      'ExpenseTrendDetailsDownload'
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
    <title>Enterprise Income / Expense Trend</title>
    <style>
        @media print {
        body {-webkit-print-color-adjust: exact;}
        }
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
            .comment {
  width: fit-content;
  background-color: #d7d7d7;
  border-radius: 0 15px 0 15px;
  padding: .2rem .5rem;
  margin-left: 2rem;
}
.bdr{
  border:1px solid white !important
 }
.notes { 
  width: fit-content;  
  padding: 0.2rem 0.5rem;
  margin-right: 3rem;
  float: inline-end;
}
   .title {
  font-family: "FaktPro-Bold";
  font-size: larger;
  padding-left: 1rem;
  text-align: left !important;
}
     .performance-scorecard-details {
          display: flex;
          flex-direction: column;
          height: auto;
          /* Adjust based on your needs */
          width: 100%;
        }
        .performance-scorecard-details .table > :not(:first-child) {
          border-top: 0px solid #ffa51a;
        }
        .performance-scorecard-details .table {
          text-align: center;
          text-transform: capitalize;
          border: transparent;
          width: 100%;
        }
        .performance-scorecard-details .table th, .performance-scorecard-details .table td {
          white-space: nowrap;
          vertical-align: top;
        }
        .performance-scorecard-details .table th:first-child, .performance-scorecard-details .table td:first-child {
          left: 0;
          z-index: 1;
        }
        .performance-scorecard-details .table tr:nth-child(odd) td:first-child, .performance-scorecard-details .table tr:nth-child(odd) td:nth-child(2) {
          background-color: #e9ecef;
        }
        .performance-scorecard-details .table tr:nth-child(even) td:first-child, .performance-scorecard-details .table tr:nth-child(even) td:nth-child(2) {
          background-color: #fff;
        }
        .performance-scorecard-details .table tr:nth-child(odd) {
          background-color: #e9ecef;
        }
        .performance-scorecard-details .table tr:nth-child(even) {
          background-color: #fff;
        }
        .performance-scorecard-details .table .spacer {
          background-color: #cfd6de !important;
          border-left: 1px solid #cfd6de !important;
          border-bottom: 1px solid #cfd6de !important;
          border-top: 1px solid #cfd6de !important;
        }
        .performance-scorecard-details .table .hidden {
          display: none !important;
        }
        .performance-scorecard-details .table .bdr-rt {
          border-right: 1px solid #abd0ec;
        }
        .performance-scorecard-details .table thead {
          position: sticky;
          top: 0;
          z-index: 99;
          font-family: 'FaktPro-Bold';
          font-size: 0.8rem;
        }
        .performance-scorecard-details .table thead th {
          padding: 5px 10px;
          margin: 0px;
          border-right: 1px solid #abd0ec;
        }
        .performance-scorecard-details .table thead .bdr-btm {
          border-bottom: #005fa3;
        }
        .performance-scorecard-details .table thead tr:nth-child(1) {
          background-color: #fff !important;
          color: #000;
          text-transform: uppercase;
          border-bottom: #cfd6de;
        }
        .performance-scorecard-details .table thead tr:nth-child(1) {
          background-color: #fff !important;
          color: #000;
          text-transform: uppercase;
          border-bottom: #cfd6de;
          box-shadow: inset 0 1px 0 0 #cfd6de;
        }
        .performance-scorecard-details .table thead tr:nth-child(1) th {
          box-shadow: inset 0 -2px 0 #337ab7;
        }
        .performance-scorecard-details .table thead tr:nth-child(3) {
          background-color: #fff !important;
          color: #000;
          text-transform: uppercase;
          border-bottom: #cfd6de;
          box-shadow: inset 0 1px 0 0 #cfd6de;
        }
        .performance-scorecard-details .table thead tr:nth-child(3) th :nth-child(1) {
          background-color: #337ab7 !important;
          color: #fff;
        }
        .performance-scorecard-details .table tbody {
          font-family: 'FaktPro-Normal';
          font-size: 0.9rem;
        }
        .performance-scorecard-details .table tbody td {
          padding: 2px 10px;
          margin: 0px;
          border: 1px solid #cfd6de;
        }
        .performance-scorecard-details .table tbody tr {
          border-bottom: 1px solid #37a6f8;
          border-left: 1px solid #37a6f8;
        }
        .performance-scorecard-details .table tbody td:first-child {
          text-align: start;
          box-shadow: inset -1px 0 0 0 #cfd6de;
        }
        .performance-scorecard-details .table tbody .sub-title {
          font-size: 0.8rem !important;
        }
        .performance-scorecard-details .table tbody .sub-subtitle {
          font-size: 0.7rem !important;
        }
        .performance-scorecard-details .table tbody .alignright {
          text-align: right;
          padding-right: 1rem;
        }
        .performance-scorecard-details .table tbody .alignleft {
          text-align: left;
          padding-left: 1rem;
        }
        .performance-scorecard-details .table tbody .text-bold {
          font-family: 'FaktPro-Bold';
        }
        .performance-scorecard-details .table tbody .darkred-bg {
          background-color: #282828 !important;
          color: #fff;
        }
        .performance-scorecard-details .table tbody .lightblue-bg {
          background-color: #646e7a !important;
          color: #fff;
        }
        .performance-scorecard-details .table tbody .gold-bg {
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
        doc.save('Enterprise Income / Expense Trend Details.pdf');
        // popupWin.close();
        this.spinner.hide();
      });
  }

  sendEmailDataDetails(Email: any, notes: any, from: any) {
    this.spinner.show();
    const printContents = document.getElementById(
      'ExpenseTrendDetailsDownload'
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
            <title>Enterprise Income / Expense Trend</title>
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
                                            .comment {
                        width: fit-content;
                        background-color: #d7d7d7;
                        border-radius: 0 15px 0 15px;
                        padding: .2rem .5rem;
                        margin-left: 2rem;
                      }
                      .bdr{
                        border:1px solid white !important
                      }
                                            .notes { 
                        width: fit-content;  
                        padding: 0.2rem 0.5rem;
                        margin-right: 3rem;
                        float: inline-end;
                      }
                              .performance-scorecard-details {
                        display: flex;
                        flex-direction: column;
                        height: auto;
                        /* Adjust based on your needs */
                        width: 100%;
                      }
                      .performance-scorecard-details .table > :not(:first-child) {
                        border-top: 0px solid #ffa51a;
                      }
                      .performance-scorecard-details .table {
                        text-align: center;
                        text-transform: capitalize;
                        border: transparent;
                        width: 100%;
                      }
                      .performance-scorecard-details .table th, .performance-scorecard-details .table td {
                        white-space: nowrap;
                        vertical-align: top;
                      }
                      .performance-scorecard-details .table th:first-child, .performance-scorecard-details .table td:first-child {
                        left: 0;
                        z-index: 1;
                      }
                      .performance-scorecard-details .table tr:nth-child(odd) td:first-child, .performance-scorecard-details .table tr:nth-child(odd) td:nth-child(2) {
                        background-color: #e9ecef;
                      }
                      .performance-scorecard-details .table tr:nth-child(even) td:first-child, .performance-scorecard-details .table tr:nth-child(even) td:nth-child(2) {
                        background-color: #fff;
                      }
                      .performance-scorecard-details .table tr:nth-child(odd) {
                        background-color: #e9ecef;
                      }
                      .performance-scorecard-details .table tr:nth-child(even) {
                        background-color: #fff;
                      }
                      .performance-scorecard-details .table .spacer {
                        background-color: #cfd6de !important;
                        border-left: 1px solid #cfd6de !important;
                        border-bottom: 1px solid #cfd6de !important;
                        border-top: 1px solid #cfd6de !important;
                      }
                      .performance-scorecard-details .table .hidden {
                        display: none !important;
                      }
                      .performance-scorecard-details .table .bdr-rt {
                        border-right: 1px solid #abd0ec;
                      }
                      .performance-scorecard-details .table thead {
                        position: sticky;
                        top: 0;
                        z-index: 99;
                        font-family: 'FaktPro-Bold';
                        font-size: 0.8rem;
                      }
                      .performance-scorecard-details .table thead th {
                        padding: 5px 10px;
                        margin: 0px;
                        border-right: 1px solid #abd0ec;
                      }
                      .performance-scorecard-details .table thead .bdr-btm {
                        border-bottom: #005fa3;
                      }
                      .performance-scorecard-details .table thead tr:nth-child(1) {
                        background-color: #fff !important;
                        color: #000;
                        text-transform: uppercase;
                        border-bottom: #cfd6de;
                      }
                      .performance-scorecard-details .table thead tr:nth-child(1) {
                        background-color: #fff !important;
                        color: #000;
                        text-transform: uppercase;
                        border-bottom: #cfd6de;
                        box-shadow: inset 0 1px 0 0 #cfd6de;
                      }
                      .performance-scorecard-details .table thead tr:nth-child(1) th {
                        box-shadow: inset 0 -2px 0 #337ab7;
                      }
                      .performance-scorecard-details .table thead tr:nth-child(3) {
                        background-color: #fff !important;
                        color: #000;
                        text-transform: uppercase;
                        border-bottom: #cfd6de;
                        box-shadow: inset 0 1px 0 0 #cfd6de;
                      }
                      .performance-scorecard-details .table thead tr:nth-child(3) th :nth-child(1) {
                        background-color: #337ab7 !important;
                        color: #fff;
                      }
                      .performance-scorecard-details .table tbody {
                        font-family: 'FaktPro-Normal';
                        font-size: 0.9rem;
                      }
                      .performance-scorecard-details .table tbody td {
                        padding: 2px 10px;
                        margin: 0px;
                        border: 1px solid #cfd6de;
                      }
                      .performance-scorecard-details .table tbody tr {
                        border-bottom: 1px solid #37a6f8;
                        border-left: 1px solid #37a6f8;
                      }
                      .performance-scorecard-details .table tbody td:first-child {
                        text-align: start;
                        box-shadow: inset -1px 0 0 0 #cfd6de;
                      }
                      .performance-scorecard-details .table tbody .sub-title {
                        font-size: 0.8rem !important;
                      }
                      .performance-scorecard-details .table tbody .sub-subtitle {
                        font-size: 0.7rem !important;
                      }
                      .performance-scorecard-details .table tbody .alignright {
                        text-align: right;
                        padding-right: 1rem;
                      }
                      .performance-scorecard-details .table tbody .alignleft {
                        text-align: left;
                        padding-left: 1rem;
                      }
                      .performance-scorecard-details .table tbody .text-bold {
                        font-family: 'FaktPro-Bold';
                      }
                      .performance-scorecard-details .table tbody .darkred-bg {
                        background-color: #282828 !important;
                        color: #fff;
                      }
                      .performance-scorecard-details .table tbody .lightblue-bg {
                        background-color: #646e7a !important;
                        color: #fff;
                      }
                      .performance-scorecard-details .table tbody .gold-bg {
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
        const pdfFile = this.blobToFile(pdfBlob, 'Enterprise Income / Expense Trend Details.pdf');
        const formData = new FormData();
        formData.append('to_email', Email);
        formData.append('subject', 'Enterprise Income / Expense Trend Details');
        formData.append('file', pdfFile);
        formData.append('notes', notes);
        formData.append('from', from);
        this.apiSrvc
          .postmethod(this.comm.routeEndpoint + 'mail', formData)
          .subscribe(
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
}
