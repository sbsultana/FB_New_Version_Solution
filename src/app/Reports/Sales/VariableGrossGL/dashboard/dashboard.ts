import { Component, HostListener, OnInit, signal } from '@angular/core';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { Api } from '../../../../Core/Providers/Api/api';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
import { BsDatepickerConfig, BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { Stores } from '../../../../CommonFilters/stores/stores';
import { Subscription } from 'rxjs';
import { NgbActiveModal } from '@ng-bootstrap/ng-bootstrap';
@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [SharedModule, BsDatepickerModule, Stores],
  templateUrl: './dashboard.html',
  styleUrls: ['./dashboard.scss']
})
export class Dashboard implements OnInit {
  SalesData: any = [];
  IndividualSalesGross: any = [];
  TotalSalesGross: any = [];
  TotalReport: any = 'B';
  TotalSortPosition: any = 'B';
  saleType: any = ['Retail'];
  Department: any = ['New', 'Used'];
  date: any = '';
  DupDate:any=''
  CurrentDate = new Date();
  NoData: boolean = false;


  stores: any = []
  groupsArray: any = [];
  storename: any = ''
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  groupId: any = 0;
  storeIds: any = 0;
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
  month!: Date;
  DuplicatDate!: Date;
  minDate!: Date;
  maxDate!: Date;
  bsConfig: Partial<BsDatepickerConfig> = {
    dateInputFormat: 'MMMM/yyyy',
    minMode: 'month'
  };

  reportOpenSub!: Subscription;
  reportGetting!: Subscription;
  Pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;
  constructor(
    public apiSrvc: Api, private ngbmodalActive: NgbActiveModal, private toast: ToastService, public shared: Sharedservice) {
    this.shared.setTitle(this.shared.common.titleName + '-Variable Gross GL');
    let today = new Date();
    let lastMonth = new Date()
    let enddate = new Date(today.setDate(today.getDate() - 1));
    this.month = new Date(enddate.setMonth(enddate.getMonth() - 1))
    this.maxDate = new Date();
    this.minDate = new Date();
    this.minDate.setFullYear(this.maxDate.getFullYear() - 3);
    this.maxDate.setMonth(this.maxDate.getMonth());
    if (today.getDate() < 5) {
      this.date = new Date(lastMonth.setMonth(lastMonth.getMonth() - 1));
    } else {
      this.date = new Date(lastMonth.setMonth(lastMonth.getMonth()));
    }
    if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
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
    this.setHeaderData()
    this.getSalesData();
  }
  ngOnInit(): void {
  }
  setHeaderData() {
    const data = {
      title: 'Variable Gross GL',
      stores: this.storeIds,
      saleType: this.saleType.toString(),
      toporbottom: this.TotalReport,
      groups: this.groupId,
      Department: this.Department,
      Month: this.date,
    };
    this.apiSrvc.SetHeaderData({ obj: data });
  }
  getSalesData() {
    if (this.storeIds != '') {
      this.shared.spinner.show();
      this.GetData();
      // this.GetTotalData();
    } else {
      this.NoData = true;
    }
  }
  Datebinding: any = ''
  GetData() {
    this.IndividualSalesGross = [];
    this.DupDate=this.date
    let date = new Date()
    let enddate = new Date(date.setDate(date.getDate() - 1));
    this.shared.datePipe.transform(this.date, 'yyyy-MM') == this.shared.datePipe.transform(enddate, 'yyyy-MM') ? this.Datebinding = 'MTD' : this.Datebinding = this.shared.datePipe.transform(this.date, 'MMM yy')
    const obj = {
      DATE: this.shared.datePipe.transform(this.date, 'yyyy-MM') + '-' + ('0' + enddate.getDate()).slice(-2),
      AS_IDS: this.storeIds.toString(),
      DEALTYPE: this.Department.indexOf('New') >= 0 && this.Department.indexOf('Used') >= 0 ? '' : this.Department.toString(),
      SALETYPE: this.saleType.toString() == 'All' ? '' : this.saleType.toString()
    };
    this.apiSrvc
      .postmethod(this.shared.common.routeEndpoint + 'GetSalesGrossGLSummaryReport', obj)
      .subscribe(
        (res) => {
          const currentTitle = document.title;
          if (res.status == 200) {
            this.IndividualSalesGross = [];
            this.TotalSalesGross = [];
            if (res.response != undefined) {
              if (res.response.length > 0) {
                let array = res.response.map((v: any) => ({
                  ...v,
                  Dealer: '-',
                }));
                let data = array.reduce(
                  (r: any, { STORE, ...rest }: any) => {
                    if (!r.some((o: any) => o.STORE == STORE)) {
                      r.push({
                        STORE,
                        ...rest,
                        subdata: array.filter(
                          (v: any) => v.STORE == STORE
                        ),
                      });
                    }
                    return r;
                  },
                  []
                );
                this.TotalSalesGross = data.filter(
                  (e: any) => e.STORE == 'Report Total'
                );
                this.IndividualSalesGross = data.filter(
                  (i: any) => i.STORE != 'Report Total'
                );
                this.SalesData = this.IndividualSalesGross
                if (this.TotalReport == 'B') {
                  this.TotalSalesGross.forEach((val: any) => {
                    this.IndividualSalesGross.push(val);
                  });
                } else {
                  this.TotalSalesGross.forEach((val: any) => {
                    this.IndividualSalesGross.unshift(val);
                  });
                }
                this.shared.spinner.hide();
              } else {
                this.shared.spinner.hide();
                this.NoData = true;
              }
            } else {
              this.shared.spinner.hide();
              this.NoData = true;
            }
          } else {
            this.toast.show(res.status, 'danger', 'Error');
            this.shared.spinner.hide();
            this.NoData = true;
          }
        },
        (error) => {
          this.toast.show('502 Bad Gate Way Error', 'danger', 'Error');
          this.shared.spinner.hide();
          this.NoData = true;
        }
      );
  }
  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    } else if (value < 0) {
      return false;
    }
    return true;
  }
  DateType: any = 'MTD';
  datetype() {
    if (this.DateType == 'PM') {
      return 'SP';
    }
    else if (this.DateType == 'C') {
      return 'C'
    }
    return this.DateType;
  }
  Salesdetails: any = []
  openDetails(Item: any, ParentItem: any, cat: any) {

    this.Salesdetails = [{
      StartDate: this.shared.datePipe.transform(this.date, 'yyyy-MM') + '-01',
      saletype: this.saleType,

      var1: Item.DEPT,
      var2: this.saleType,
      var3: '',

      var1Value: Item.data1,
      var2Value: '',
      var3Value: '',

      userName: Item.STORE,
      storeids: Item.STOREID == 0 ? this.storeIds : Item.STOREID,

      DepartmentN: this.Department.includes('New') ? 'N' : '',
      DepartmentU: this.Department.includes('Used') ? 'U' : '',

      parent: ParentItem,
      category: cat
    }];
    this.expandedIndex = null;
    this.FSSubDetailsMap = {};
    this.currentPage = 1;
    this.GetDetails('', 0);
  }
  details: any = [];
  spinnerLoader: boolean = true;
  DetailsSearchName: any;
  SubDetailsSearchName: any;
  acctNo: any = '';
  SalesdetailsData: any = [];
  filteredSalesdetailsData: any[] = [];
  searchText: string = '';
  ETdetailsData: any = [];
  currentPage: number = 1;
  itemsPerPage: number = 100;
  maxPageButtonsToShow: number = 3;
  clickedPage: number | null = null;
  Opacity: any = 'N';


  GetDetails(acctno: any, index: number) {

    // 🔹 MAIN DATA LOAD
    if (!acctno) {
      this.spinnerLoader = true;

      const obj = {
        AS_IDS: this.Salesdetails[0].storeids,
        DATE: this.Salesdetails[0].StartDate,
        VAR1: this.Salesdetails[0].var1,
        VAR2: this.Salesdetails[0].var2.toString(),
        Accountnumber: ''
      };

      this.apiSrvc
        .postmethod(this.shared.common.routeEndpoint + 'GetSalesGrossGLDetailsV1', obj)
        .subscribe((res: any) => {

          if (res.status === 200) {

            this.SalesdetailsData = (res.response || []).map((item: any) => ({
              Store: item.Store || item.As_dealername || '-',
              AccountNumber: item.AccountNumber || '-',
              AccountDescription: item['Account Description'] || '-',
              Balance: item.Balance || 0
            }));

            this.filterData();
            this.spinnerLoader = false;
            this.NoData = this.SalesdetailsData.length === 0;
          }
        });

      return;
    }

    // 🔹 TOGGLE CLOSE (if already open)
    if (this.expandedIndex === index) {
      this.expandedIndex = null;
      return;
    }

    // 🔹 IF DATA ALREADY LOADED → JUST OPEN (no API call again)
    if (this.FSSubDetailsMap[index]) {
      this.expandedIndex = index;
      return;
    }

    // 🔹 LOAD SUB DETAILS
    this.spinnerLoader = true;

    const obj = {
      AS_IDS: this.Salesdetails[0].storeids,
      DATE: this.Salesdetails[0].StartDate,
      VAR1: this.Salesdetails[0].var1,
      VAR2: this.Salesdetails[0].var2.toString(),
      Accountnumber: acctno
    };

    this.apiSrvc
      .postmethod(this.shared.common.routeEndpoint + 'GetSalesGrossGLDetailsV1', obj)
      .subscribe((res: any) => {

        if (res.status === 200) {

          this.FSSubDetailsMap[index] = (res.response || []).map((item: any) => ({
            Control: item.Control || '-',
            Date: item.Date,
            AccountDescription: item['Account Description'] || '-',
            Balance: item.Balance || 0
          }));

          this.expandedIndex = index;
          this.spinnerLoader = false;
        }
      });
  }
  expandedIndex: number | null = null;
  FSSubDetailsMap: { [index: number]: any[] } = {};


  get postingAmountTotal(): number {
    return this.filteredSalesdetailsData.reduce((total, item) => {
      return total + (item.Balance || 0);
    }, 0);
  }
  getPostingSubAmountTotal(index: number): number {
    const subRows = this.FSSubDetailsMap[index];
    if (!subRows || !Array.isArray(subRows)) {
      return 0;
    }

    return subRows.reduce((total, item) => {
      return total + (item.Balance || 0);
    }, 0);
  }

  filterData() {
    const text = this.searchText.trim().toLowerCase();

    if (!text) {
      this.filteredSalesdetailsData = [...this.SalesdetailsData];
    } else {
      this.filteredSalesdetailsData = this.SalesdetailsData.filter((item: any) =>
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
    return this.filteredSalesdetailsData.slice(start, start + this.itemsPerPage);
  }

  getMaxPageNumber(): number {
    return Math.max(1,
      Math.ceil(this.filteredSalesdetailsData.length / this.itemsPerPage)
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
      this.filteredSalesdetailsData.length
    );
  }
  onPageSizeChange() {
    this.currentPage = 1;
  }

  resetExpand() {
    this.clickedPage = null;
    this.expandedIndex = null;
  }


  close() {
    this.ngbmodalActive.close();
    console.log(this.Opacity);
    this.filteredSalesdetailsData = [];
    if (this.Salesdetails.STORES != '') {
      this.goToFirstPage();
    }
  }
  onclose() {
    this.shared.ngbmodal.dismissAll();
    console.log(this.Opacity);
  }

  Account_Details: any = [];
  AcctDetails: any = [];
  Acct_ID: any;
  Obj: any;



  isDesc: boolean = false;
  column: string = 'CategoryName';
  sort(property: any, data: any) {
    this.isDesc = !this.isDesc; //change the direction
    this.column = property;
    let direction = this.isDesc ? 1 : -1;
    // //console.log(direction);
    if (direction == -1) {
      this.TotalSortPosition = 'T';
    } else {
      this.TotalSortPosition = 'B';
    }
    data.sort(function (a: any, b: any) {
      if (a[property] < b[property]) {
        return -1 * direction;
      } else if (a[property] > b[property]) {
        return 1 * direction;
      } else {
        return 0;
      }
    });
  }
  subdataindex: any = 0;
  ngAfterViewInit() {
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.shared.common.pageName == 'Variable Gross GL') {
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
        if (res.obj.Module == 'Variable Gross GL') {
          document.getElementById('report')?.click();
        }
      }
    });
    this.excel = this.apiSrvc.getExportToExcelAllReports().subscribe((res) => {
      if (this.excel != undefined) {
        if (res.obj.title == 'Variable Gross GL') {
          if (res.obj.state == true) {
            this.exportToExcel();
          }
        }
      }
    });
    this.email = this.apiSrvc.getExportToEmailPDFAllReports().subscribe((res) => {
      if (this.email != undefined) {
        if (res.obj.title == 'Variable Gross GL') {
          if (res.obj.stateEmailPdf == true) {
            // this.sendEmailData(res.obj.Email, res.obj.notes, res.obj.from);
          }
        }
      }
    });
    this.print = this.apiSrvc.getExportToPrintAllReports().subscribe((res) => {
      if (this.print != undefined) {
        if (res.obj.title == 'Variable Gross GL') {
          if (res.obj.statePrint == true) {
            // this.GetPrintData();
          }
        }
      }
    });
    this.Pdf = this.apiSrvc.getExportToPDFAllReports().subscribe((res) => {
      if (this.Pdf != undefined) {
        if (res.obj.title == 'Variable Gross GL') {
          if (res.obj.statePDF == true) {
            // this.generatePDF();
          }
        }
      }
    });
  }
  StoresData(data: any) {
    this.storeIds = data.storeids;
    this.groupId = data.groupId;
    this.storename = data.storename;
    this.groupName = data.groupName;
    this.storecount = data.storecount;
    this.storedisplayname = data.storedisplayname;
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
  }


  onOpenCalendar(container: any) {
    container.setViewMode('month');
    container.monthSelectHandler = (event: any): void => {
      container.value = event.date;
      this.date = event.date;
      return;
    };
  }
  changeDate(e: any) {
    console.log(e);
    this.date = e;
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

  multipleorsingle(block: any, e: any) {

    if (block == 'RL') {
      this.saleType = []
      this.saleType.push(e);
    }

    if (block == 'TB') {
      this.TotalReport = e;

    }

    if (block == 'Dept') {
      const index = this.Department.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.Department.splice(index, 1);
        if (this.Department.length == 0) {
          this.toast.show('Please Select Atleast One Department', 'warning', 'Warning');
        }
      } else {
        this.Department.push(e);
      }
    }
  }
  activePopover: number = -1;
  togglePopover(popoverIndex: number) {

    if (this.activePopover === popoverIndex) {
      // If the same popover is clicked, close it
      this.activePopover = -1;
    } else {
      // Open the selected popover and close others
      this.activePopover = popoverIndex;
    }
  }
  viewreport() {

    this.activePopover = -1
    if (this.storeIds.length == 0 || this.Department.length == 0) {
      if (this.Department.length == 0) {
        this.toast.show('Please Select Atleast One Department', 'warning', 'Warning');
      }
      if (this.storeIds.length == 0) {
        this.toast.show('Please Select Atleast One Store', 'warning', 'Warning');
      }
    } else {
      this.setHeaderData()
      this.getSalesData()
    }
    // }
    // }
  }
  Scrollpercent: any = 0;
  updateVerticalScroll(event: any): void {
    const scrollDemo = document.querySelector('#scrollcent');
    this.Scrollpercent = Math.round(
      (event.target.scrollTop /
        (event.target.scrollHeight - scrollDemo!.clientHeight)) *
      100
    );
  }
  index = '';
  ExcelStoreNames: any = [];
  exportToExcel() {
    let storeNames: any = [];
    let store = this.storeIds.split(',');
    storeNames = this.shared.common.groupsandstores
      .filter((v: any) => v.sg_id == this.groupId)[0]
      .Stores.filter((item: any) =>
        store.some((cat: any) => cat === item.ID.toString())
      );
    if (
      store.length ==
      this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0]
        .Stores.length
    ) {
      this.ExcelStoreNames = 'All Stores';
    } else {
      this.ExcelStoreNames = storeNames.map(function (a: any) {
        return a.storename;
      });
    }
    const SalesData = this.SalesData.map((_arrayElement: any) =>
      Object.assign({}, _arrayElement)
    );
    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('Variable Gross GL');
    worksheet.views = [
      {
        // state: 'frozen',
        // ySplit: 24, // Number of rows to freeze (2 means the first two rows are frozen)
        // topLeftCell: 'A25', // Specify the cell to start freezing from (in this case, the third row)
        showGridLines: false,
      },
    ];
    worksheet.addRow('');
    const titleRow = worksheet.addRow(['Variable Gross GL']);
    titleRow.eachCell((cell: any, number: any) => {
      cell.alignment = { indent: 1, vertical: 'middle', horizontal: 'left' };
    });
    titleRow.font = { name: 'Arial', family: 4, size: 12, bold: true };
    titleRow.worksheet.mergeCells('A2', 'D2');
    worksheet.addRow('');
    const PresentMonth = this.shared.datePipe.transform(this.date, 'MMMM yyyy');
    const DateToday = this.shared.datePipe.transform(
      new Date(),
      'MM/dd/yyyy h:mm:ss a'
    );
    const DATE_EXTENSION = this.shared.datePipe.transform(new Date(), 'MMddyyyy');
    worksheet.addRow([DateToday]).font = { name: 'Arial', family: 4, size: 9 };
    const ReportFilter = worksheet.addRow(['Report Controls :']);
    ReportFilter.font = { name: 'Arial', family: 4, size: 10, bold: true };
    // const Groupings = worksheet.addRow(['Groupings :']);
    // Groupings.getCell(1).font = {
    //   name: 'Arial',
    //   family: 4,
    //   size: 9,
    //   bold: true,
    // };
    // const groupings = worksheet.getCell('B6');
    // groupings.value = this.path1name + ((this.path2name != '' && this.path2name != undefined) ? (', ' + this.path2name) : '') + ((this.path3name != '' && this.path3name != undefined) ? (', ' + this.path3name) : '');
    // groupings.font = { name: 'Arial', family: 4, size: 9 };
    const Timeframe = worksheet.addRow(['Timeframe :']);
    Timeframe.getCell(1).font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
    };
    const timeframe = worksheet.getCell('B6');
    timeframe.value = PresentMonth
    timeframe.font = { name: 'Arial', family: 4, size: 9 };
    const Stores = worksheet.addRow(['Stores :']);
    Stores.getCell(1).font = { name: 'Arial', family: 4, size: 9, bold: true };
    const Groups = worksheet.getCell('B8');
    Groups.value = 'Group :';
    const groups = worksheet.getCell('D8');
    groups.value = this.shared.common.groupsandstores.filter(
      (val: any) => val.sg_id == this.groupId.toString()
    )[0].sg_name;
    groups.font = { name: 'Arial', family: 4, size: 9 };
    const Brands = worksheet.getCell('B9');
    Brands.value = 'Brands :';
    const brands = worksheet.getCell('D9');
    brands.value = '-';
    brands.font = { name: 'Arial', family: 4, size: 9 };
    const Stores1 = worksheet.getCell('B10');
    Stores1.value = 'Stores :';
    worksheet.mergeCells('D10', 'O12');
    const stores1 = worksheet.getCell('D10');
    stores1.value =
      this.ExcelStoreNames == 0
        ? 'All Stores'
        : this.ExcelStoreNames == null
          ? '-'
          : this.ExcelStoreNames.toString().replaceAll(',', ', ');
    stores1.font = { name: 'Arial', family: 4, size: 9 };
    stores1.alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    const Filters = worksheet.addRow(['Filters :']);
    Filters.getCell(1).font = { name: 'Arial', family: 4, size: 9, bold: true };
    const NewUsed = worksheet.getCell('B16');
    NewUsed.value = 'Department :';
    const newused = worksheet.getCell('D16');
    newused.value = this.Department.toString().replaceAll(',', ', ');
    newused.font = { name: 'Arial', family: 4, size: 9 };
    worksheet.addRow('');
    let dateYear = worksheet.getCell('A18');
    dateYear.value = PresentMonth;
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
      fgColor: { argb: '2a91f0' },
      bgColor: { argb: 'FF0000FF' },
    };
    dateYear.border = { right: { style: 'dotted' } };
    worksheet.mergeCells('B18', 'I18');
    let units = worksheet.getCell('B18');
    units.value = 'Units';
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
      fgColor: { argb: '2a91f0' },
      bgColor: { argb: 'FF0000FF' },
    };
    units.border = { right: { style: 'dotted' } };
    worksheet.mergeCells('J18', 'T18');
    let frontgross = worksheet.getCell('J18');
    frontgross.value = 'Gross';
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
      fgColor: { argb: '2a91f0' },
      bgColor: { argb: 'FF0000FF' },
    };
    frontgross.border = { right: { style: 'dotted' } };
    let Headings = [
      '',
      'MTD',
      'Pace',
      'Target',
      'Diff',
      'LY',
      'YOY%',
      'LM',
      'MOM%',
      'MTD',
      'Front Gross',
      'Back Gross',
      'PVR',
      'Pace',
      'Target',
      'Diff',
      'LY',
      'YOY%',
      'LM',
      'MOM%',
    ];
    const headerRow = worksheet.addRow(Headings);
    headerRow.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: false,
      color: { argb: 'FFFFFF' },
    };
    headerRow.alignment = {
      indent: 1,
      vertical: 'middle',
      horizontal: 'center',
    };
    headerRow.eachCell((cell: any, number: any) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '788494' },
        bgColor: { argb: 'FF0000FF' },
      };
      cell.border = { right: { style: 'dotted' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });
    for (const d of SalesData) {
      const obj = [
        d.STORE == '' ? '-' : d.STORE == null ? '-' : d.STORE,
        d.UNITS_MTD == '' ? '-' : d.UNITS_MTD == null ? '-' : d.UNITS_MTD,
        d.UNITS_PACE == '' ? '-' : d.UNITS_PACE == null ? '-' : d.UNITS_PACE,
        d.UNITS_TARGET == ''
          ? '-'
          : d.UNITS_TARGET == null
            ? '-'
            : d.UNITS_TARGET,
        d.UNITS_DIFF == '' ? '-' : d.UNITS_DIFF == null ? '-' : d.UNITS_DIFF,
        d.UNITS_LY == '' ? '-' : d.UNITS_LY == null ? '-' : d.UNITS_LY,
        d.UNITS_YOY == ''
          ? '-'
          : d.UNITS_YOY == null
            ? '-'
            : parseFloat(d.UNITS_YOY) + '%',
        d.UNITS_LM == '' ? '-' : d.UNITS_LM == null ? '-' : d.UNITS_LM,
        d.UNITS_MOM == '' ? '-' : d.UNITS_MOM == null ? '-' : d.UNITS_MOM + '%',
        d.GROSS_MTD == '' ? '-' : d.GROSS_MTD == null ? '-' : d.GROSS_MTD,
        d.GROSS_FG == '' ? '-' : d.GROSS_FG == null ? '-' : d.GROSS_FG,
        d.GROSS_BG == '' ? '-' : d.GROSS_BG == null ? '-' : d.GROSS_BG,
        d.GROSS_PVR == ''
          ? '-'
          : d.GROSS_PVR == null
            ? '-'
            : parseFloat(d.GROSS_PVR),
        d.GROSS_PACE == ''
          ? '-'
          : d.GROSS_PACE == null
            ? '-'
            : parseFloat(d.GROSS_PACE),
        d.GROSS_TARGET == ''
          ? '-'
          : d.GROSS_TARGET == null
            ? '-'
            : d.GROSS_TARGET,
        d.GROSS_DIFF == '' ? '-' : d.GROSS_DIFF == null ? '-' : d.GROSS_DIFF,
        d.GROSS_LY == '' ? '-' : d.GROSS_LY == null ? '-' : d.GROSS_LY,
        d.GROSS_YOY == ''
          ? '-'
          : d.GROSS_YOY == null
            ? '-'
            : parseFloat(d.GROSS_YOY) + '%',
        d.GROSS_LM == ''
          ? '-'
          : d.GROSS_LM == null
            ? '-'
            : parseFloat(d.GROSS_LM),
        d.GROSS_MOM == '' ? '-' : d.GROSS_MOM == null ? '-' : d.GROSS_MOM + '%',
      ];
      const Data1 = worksheet.addRow(obj);
      Data1.font = { name: 'Arial', family: 4, size: 9 };
      Data1.alignment = { vertical: 'middle', horizontal: 'center' };
      Data1.getCell(1).alignment = {
        indent: 1,
        vertical: 'middle',
        horizontal: 'left',
      };
      Data1.eachCell((cell: any, number: any) => {
        cell.border = { right: { style: 'dotted' } };
        cell.numFmt = '$#,##0.00';
        if (number > 1 && number < 10) {
          cell.numFmt = '#,##0';
        }
        if (number >= 10) {
          cell.numFmt = '$#,##0';
        }
        if (number > 1 && obj[number] != undefined) {
          if (obj[number] < 0) {
            Data1.getCell(number + 1).font = {
              name: 'Arial',
              family: 4,
              size: 9,
              color: { argb: 'FFFF0000' },
            };
          }
        }
      });
      if (Data1.number % 2) {
        Data1.eachCell((cell, number) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'e5e5e5' },
            bgColor: { argb: 'FF0000FF' },
          };
        });
      }
      if (d.subdata != undefined) {
        // for (const d1 of d.subdata)
        d.subdata.forEach((d1: any, i: any) => {
          if (i != 0) {
            const obj = [
              d1.DEPT == '' ? '-' : d1.DEPT == null ? '-' : d1.DEPT,
              d1.UNITS_MTD == ''
                ? '-'
                : d1.UNITS_MTD == null
                  ? '-'
                  : d1.UNITS_MTD,
              d1.UNITS_PACE == ''
                ? '-'
                : d1.UNITS_PACE == null
                  ? '-'
                  : d1.UNITS_PACE,
              d1.UNITS_TARGET == ''
                ? '-'
                : d1.UNITS_TARGET == null
                  ? '-'
                  : d1.UNITS_TARGET,
              d1.UNITS_DIFF == ''
                ? '-'
                : d1.UNITS_DIFF == null
                  ? '-'
                  : d1.UNITS_DIFF,
              d1.UNITS_LY == '' ? '-' : d1.UNITS_LY == null ? '-' : d1.UNITS_LY,
              d1.UNITS_YOY == ''
                ? '-'
                : d1.UNITS_YOY == null
                  ? '-'
                  : parseFloat(d1.UNITS_YOY),
              d1.UNITS_LM == '' ? '-' : d1.UNITS_LM == null ? '-' : d1.UNITS_LM,
              d1.UNITS_MOM == ''
                ? '-'
                : d1.UNITS_MOM == null
                  ? '-'
                  : d1.UNITS_MOM,
              d1.GROSS_MTD == ''
                ? '-'
                : d1.GROSS_MTD == null
                  ? '-'
                  : d1.GROSS_MTD,
              d1.GROSS_FG == '' ? '-' : d1.GROSS_FG == null ? '-' : d1.GROSS_FG,
              d1.GROSS_BG == '' ? '-' : d1.GROSS_BG == null ? '-' : d1.GROSS_BG,
              d1.GROSS_PVR == ''
                ? '-'
                : d1.GROSS_PVR == null
                  ? '-'
                  : parseFloat(d1.GROSS_PVR),
              d1.GROSS_PACE == ''
                ? '-'
                : d1.GROSS_PACE == null
                  ? '-'
                  : parseFloat(d1.GROSS_PACE),
              d1.GROSS_TARGET == ''
                ? '-'
                : d1.GROSS_TARGET == null
                  ? '-'
                  : d1.GROSS_TARGET,
              d1.GROSS_DIFF == ''
                ? '-'
                : d1.GROSS_DIFF == null
                  ? '-'
                  : d1.GROSS_DIFF,
              d1.GROSS_LY == '' ? '-' : d1.GROSS_LY == null ? '-' : d1.GROSS_LY,
              d1.GROSS_YOY == ''
                ? '-'
                : d1.GROSS_YOY == null
                  ? '-'
                  : parseFloat(d1.GROSS_YOY),
              d1.GROSS_LM == ''
                ? '-'
                : d1.GROSS_LM == null
                  ? '-'
                  : parseFloat(d1.GROSS_LM),
              d1.GROSS_MOM == ''
                ? '-'
                : d1.GROSS_MOM == null
                  ? '-'
                  : d1.GROSS_MOM,
            ];
            const Data2 = worksheet.addRow(obj);
            Data2.outlineLevel = 1; // Grouping level 2
            Data2.font = { name: 'Arial', family: 4, size: 9 };
            Data2.alignment = { vertical: 'middle', horizontal: 'center' };
            Data2.getCell(1).alignment = {
              indent: 2,
              vertical: 'middle',
              horizontal: 'left',
            };
            Data2.eachCell((cell: any, number: any) => {
              cell.border = { right: { style: 'dotted' } };
              cell.numFmt = '$#,##0';
              if (number > 1 && number < 10) {
                cell.numFmt = '#,##0';
              }
              if (number >= 10) {
                cell.numFmt = '$#,##0';
              }
              if (number > 1 && obj[number] != undefined) {
                if (obj[number] < 0) {
                  Data1.getCell(number + 1).font = {
                    name: 'Arial',
                    family: 4,
                    size: 9,
                    color: { argb: 'FFFF0000' },
                  };
                }
              }
            });
            if (Data2.number % 2) {
              Data2.eachCell((cell, number) => {
                cell.fill = {
                  type: 'pattern',
                  pattern: 'solid',
                  fgColor: { argb: 'e5e5e5' },
                  bgColor: { argb: 'FF0000FF' },
                };
              });
            }
          }
        });
        //  {
        // }
      }
      if (d.data1 === 'Report Total') {
        Data1.eachCell((cell) => {
          cell.font = { name: 'Arial', family: 4, size: 9, bold: true };
          cell.border = {
            top: { style: 'thin' },
            bottom: { style: 'thin' },
          };
        });
      }
    }
    worksheet.eachRow((row, rowIndex) => {
      row.eachCell((cell, colIndex) => {
        if (rowIndex > 1 && rowIndex < 21) {
          // Skip the header row
          // Apply conditional alignment based on your conditions
          if (colIndex === 1) {
            // Apply right alignment to the second column
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
    worksheet.getColumn(2).width = 12;
    worksheet.getColumn(3).width = 12;
    worksheet.getColumn(4).width = 12;
    worksheet.getColumn(5).width = 12;
    worksheet.getColumn(6).width = 12;
    worksheet.getColumn(7).width = 12;
    worksheet.getColumn(8).width = 12;
    worksheet.getColumn(9).width = 12;
    worksheet.getColumn(10).width = 12;
    worksheet.getColumn(11).width = 12;
    worksheet.getColumn(12).width = 12;
    worksheet.getColumn(13).width = 12;
    worksheet.getColumn(14).width = 12;
    worksheet.getColumn(15).width = 12;
    worksheet.getColumn(16).width = 12;
    worksheet.getColumn(17).width = 12;
    worksheet.getColumn(18).width = 12;
    worksheet.getColumn(19).width = 12;
    worksheet.getColumn(20).width = 12;
    worksheet.getColumn(21).width = 12;
    worksheet.addRow([]);
    workbook.xlsx.writeBuffer().then(buffer => {
      this.shared.exportToExcel(workbook, 'Variable Gross GL');
    });
  }


  AcctDetasil_ExportAsXLSX() {
    let localarray = this.details.map((_arrayElement: any) =>
      Object.assign({}, _arrayElement)
    );
    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('Variable Gross GL');
    worksheet.views = [
      {
        state: 'frozen',
        ySplit: 5, // Number of rows to freeze (2 means the first two rows are frozen)
        topLeftCell: 'A13', // Specify the cell to start freezing from (in this case, the third row)
        showGridLines: false,
      },
    ];
    worksheet.addRow('')

    const DateToday = this.shared.datePipe.transform(new Date(), 'MM/dd/yyyy h:mm:ss a');

    const titleRow = worksheet.getCell("A2"); titleRow.value = 'Variable Gross GL';
    titleRow.font = { name: 'Arial', family: 4, size: 15, bold: true };
    titleRow.alignment = { indent: 1, vertical: 'middle', horizontal: 'left' }



    const DateBlock = worksheet.getCell("L2"); DateBlock.value = DateToday;
    DateBlock.font = { name: 'Arial', family: 4, size: 10 };
    DateBlock.alignment = { vertical: 'middle', horizontal: 'center' }
    worksheet.addRow([''])




    const DATE_EXTENSION = this.shared.datePipe.transform(
      new Date(),
      'MMddyyyy'
    );



    worksheet.addRow('')
    let Headings = [
      'Sl.no',
      'Store',
      'Account #',
      'Description',
      'Balance',
      'Control',
      'Date',
    ];


    const headerRow = worksheet.addRow(Headings);
    headerRow.font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFFFF' }, }
    headerRow.alignment = { indent: 1, vertical: 'middle', horizontal: 'center' };
    headerRow.height = 22;
    headerRow.eachCell((cell, number) => {
      cell.fill = {
        type: 'pattern', pattern: 'solid', fgColor: { argb: '2a91f0' }, bgColor: { argb: 'FF0000FF' }
      }
      cell.border = { right: { style: 'thin' } }
      cell.alignment = { vertical: 'middle', horizontal: 'center' }
    });

    // //console.log(localarray);
    var count = 0
    for (const d of localarray) {
      count++
      d.contractdate = this.shared.datePipe.transform(d.Date, 'MM.dd.yyyy');

      var obj = [count,
        (d['As_dealername'] == '' ? '-' : (d['As_dealername'] == null ? '-' : (d['As_dealername']))),
        (d['ACCOUNT#'] == '' ? '-' : (d['ACCOUNT#'] == null ? '-' : (d['ACCOUNT#']))),
        (d['Account Description'] == '' ? '-' : (d['Account Description'] == null ? '-' : (d['Account Description']))),
        (d.Balance == '' ? '-' : (d.Balance == null ? '-' : (parseFloat(d.Balance)))),
        (d.Control == '' ? '-' : (d.Control == null ? '-' : d.Control)),
        (d.contractdate == '' ? '-' : (d.contractdate == null ? '-' : (d.contractdate))),


      ];


      const Data1 = worksheet.addRow(obj);

      // //console.log(Data1);

      Data1.font = { name: 'Arial', family: 4, size: 8 }
      Data1.alignment = { vertical: 'middle', horizontal: 'center', indent: 1 }
      // Data1.getCell(1).alignment = {indent: 1,vertical: 'top', horizontal: 'left'}
      Data1.eachCell((cell, number) => {
        cell.border = { right: { style: 'thin' } }
        if (number == 5) {
          cell.numFmt = '$#,##0'

        }

      });
      if (Data1.number % 2) {
        Data1.eachCell((cell, number) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'e5e5e5' },
            bgColor: { argb: 'FF0000FF' },
          };
        });
      }
      Data1.worksheet.columns.forEach((column: any, columnIndex: any) => {
        let maxLength = 0;
        column.eachCell({ includeEmpty: true }, (cell: any) => {
          const length = cell.value ? cell.value.toString().length : 10;
          if (length > maxLength) {
            maxLength = length;
          }
        });
        column.width = maxLength < 10 ? 10 : maxLength + 2; // Set a minimum width of 10
      });

      // });
      // count++
    }
    worksheet.getColumn(1).width = 16;
    worksheet.getColumn(2).width = 16;
    worksheet.getColumn(3).width = 20;
    worksheet.getColumn(4).width = 20;
    worksheet.getColumn(5).width = 20;



    worksheet.addRow([]);


    workbook.xlsx.writeBuffer().then(buffer => {
      this.shared.exportToExcel(workbook, 'Variable Gross GL Account Level Details');
    });
  }
}