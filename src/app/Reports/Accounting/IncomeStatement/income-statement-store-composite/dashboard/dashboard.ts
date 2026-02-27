import { Component, Input, OnInit, ViewChild, ElementRef, HostListener } from '@angular/core';
import { BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { Subscription } from 'rxjs';
import { BsDatepickerConfig } from 'ngx-bootstrap/datepicker';
import { SharedModule } from '../../../../../Core/Providers/Shared/shared.module';
import { Sharedservice } from '../../../../../Core/Providers/Shared/sharedservice';
import { Setdates } from '../../../../../Core/Providers/SetDates/setdates';
import { common } from '../../../../../common';
import { formatDate } from '@angular/common'
@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [SharedModule, BsDatepickerModule],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss'
})
export class Dashboard {

  Current_Date: any;


  FSData: any = [];
  StoreNamesHeadings: any = [];

  NoData: boolean = false;
  ExpenseTrendByStoreKeys: any = [];
  AllDatakeys: any = [];
  ExpenseTrendByStore_Excel: any;
  XpenseTrendByStoreKeys: string[] = [];
  ExpenseTrendByStore: any;
  ExpenseTrendByStore_ExcelMonth: any;

  Month: any;

  newdate: Date = new Date();
  date: any = new Date(this.newdate.setMonth(this.newdate.getMonth() - 1));


  constructor(
    public shared: Sharedservice, public setdates: Setdates, private comm: common,
  ) {

    const lastMonth = new Date();
    let today = new Date();
    if (today.getDate() < 5) {
      this.date = new Date(lastMonth.setMonth(lastMonth.getMonth() - 1));
    } else {
      this.date = new Date(lastMonth.setMonth(lastMonth.getMonth()));
    }
    this.Month = new Date(this.date)
    // ('0' + (this.date.getMonth() + 1)).slice(-2) +
    // '-' +
    // ('0' + this.date.getDate()).slice(-2) +
    // '-' +
    // this.date.getFullYear();
    // this.Month='01-12-2022'
    // this.Month= new Date(newdate.getFullYear(), newdate.getMonth(), 1)
    this.shared.setTitle(this.shared.common.titleName + '-Income Statement Store Composite');

    // if (localStorage.getItem('Fav') != 'Y') {
    // this.groups= this.comm.TokenGroups;
    const data = {
      title: 'Income Statement Store Composite',
      Month: this.date,
      stores: 2,
      groups: 1,
    };
    this.shared.api.SetHeaderData({
      obj: data,
    });





    this.GetData();
    // } else {
    //   //alert('Hello Fav..!')
    //   // this.getFavReports();
    // }
    window.addEventListener('scroll', function () {
      const maxHeight = document.body.scrollHeight - window.innerHeight;
      // //console.log((window.pageYOffset * 100) / maxHeight);
    });
  }
  bsConfig: Partial<BsDatepickerConfig> = {
    dateInputFormat: 'MMMM/YYYY',
    minMode: 'month'
  };
  ngOnInit(): void {

    // this.GetDetails()
    // var curl =
    //   this.comm.menuUrl+'/favouritereports/GetIncomeStatementStoreComposite';
    // this.apiSrvc.logSaving(
    //   curl,
    //   {},
    //   '',
    //   'Success',
    //   'Income Statement Store Composite'
    // );
    this.maxDate = new Date();
    this.minDate = new Date();
    this.minDate.setFullYear(this.maxDate.getFullYear() - 3);
    this.maxDate.setDate(this.maxDate.getDate());
  }
  Scrollpercent: any = 0;
  updateVerticalScroll(event: any): void {
    const scrollDemo = document.querySelector('#scrollcent') as HTMLElement;
    this.Scrollpercent = Math.round(
      (event.target.scrollTop /
        (event.target.scrollHeight - scrollDemo.clientHeight)) *
      100
    );

    // //console.log(
    //   event.target.scrollTop,
    //   event.document,
    //   ((document.documentElement.scrollTop + document.body.scrollTop) /
    //     (document.documentElement.scrollHeight -
    //       document.documentElement.clientHeight)) *
    //     100
    // );

    // var el = document.getElementById('scrollcent');
    // var top = el.offsetTop;
    // var height = el.offsetHeight;
    // var scrollTop = document.body.scrollTop;
    // var difference = el.scrollTop - top;

    // var percent = difference / height;
    // //console.log(el.scrollTop, el.offsetHeight, percent, scrollTop);
  }
  GetData() {
    ////console.log(this.Month)
    this.shared.spinner.show();
    this.ExpenseTrendByStoreKeys = [];
    this.AllDatakeys = [];
    //   let date =
    //   ('0' + (this.month.getMonth() + 1)).slice(-2) +
    //   '-01' +
    //   '-' +
    //   this.month.getFullYear();
    const obj = {
      // "SalesDate": "11-01-2022"
      SalesDate: this.shared.datePipe.transform(this.Month, 'MM-dd-yyyy'),
      Stores: '2',
      UserID: 0,
      // UserID: 1,
    };
    let startFrom = new Date().getTime();
    const curl = this.shared.getEnviUrl() + this.comm.routeEndpoint + 'GetIncomeStatementStoreComposite';
    this.shared.api
      .postmethod(this.comm.routeEndpoint + 'GetIncomeStatementStoreComposite', obj)
      .subscribe(
        (x: { message: any; status: number; response: string | any[] | undefined; }) => {
          const currentTitle = document.title;
          this.shared.api.logSaving(curl, {}, '', x.message, currentTitle);
          if (x.status == 200) {
            if (x.response != undefined) {
              if (x.response.length > 0) {
                let resTime = (new Date().getTime() - startFrom) / 1000;
                // this.logSaving(
                //   this.solutionurl + 'Price/GetIncomeStatementStoreComposite',
                //   obj,
                //   resTime,
                //   'Success'
                // );

                ////console.log('Overall Response',x.response)
                // this.ExpenseTrendByStore_Excel = x.response.map(
                //   ({ SNo, ...rest }) => ({ ...rest })
                // );
                ////console.log('ExpenseTrendByStore_Excel',this.ExpenseTrendByStore_Excel)
                this.XpenseTrendByStoreKeys = Object.keys(x.response[0]);
                this.Fsdetails = Object.keys(x.response[0]);
                // //console.log('XpenseTrendByStoreKeys', this.XpenseTrendByStoreKeys);
                let ETByStoreKeys_sorted = this.XpenseTrendByStoreKeys.filter(
                  (x) => {
                    return x != 'SNo' && x != 'NgClass';
                  }
                );

                // let ETByStoreKeys_sorted = this.XpenseTrendByStoreKeys.filter(x => { return (x != "" && x != "" && x != "") }).sort((a,b) => a.localeCompare(b));
                ////console.log('ETByStoreKeys_sorted',ETByStoreKeys_sorted)
                let AllStore_Label = ETByStoreKeys_sorted.find((x: any) => {
                  if (x.toUpperCase() == 'All Stores'.toUpperCase()) return x;
                });
                // //console.log('AllStore_Label',AllStore_Label)
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
                const Cmtindex = ETByStoreKeys_sorted.findIndex(
                  (i) => i == 'CommentsStatus'
                );
                if (Cmtindex >= 0) {
                  ETByStoreKeys_sorted.splice(Cmtindex, 1);
                  // ETByStoreKeys_sorted.splice(0, 0, 'CommentsStatus');
                }
                //console.log('ETByStoreKeys_sorted', ETByStoreKeys_sorted);
                // /* To add space between each array value to get dynamic <th>*/
                // this.AllDatakeys= ["LABLE1", "NgClass"];
                for (var i = 0; i < ETByStoreKeys_sorted.length; i++) {
                  this.ExpenseTrendByStoreKeys.push(
                    ETByStoreKeys_sorted[i].toLowerCase()
                  );
                  // this.ExpenseTrendByStoreKeys.push("");

                  this.AllDatakeys.push(ETByStoreKeys_sorted[i]);

                  // this.AllDatakeys.push("");
                }
                this.ExpenseTrendByStoreKeys.push(AllStore_Label);
                this.AllDatakeys.push(AllStore_Label);
                // //console.log('ExpenseTrendByStoreKeys',this.ExpenseTrendByStoreKeys)
                // //console.log('AllDatakeys', this.AllDatakeys);
                let XpenseTrendByStoreData = x.response;
                let ETByStoreData_sorted = XpenseTrendByStoreData;
                let AllStoreLabel_Data = '';
                this.ExpenseTrendByStore = x.response;
                ////console.log('ExpenseTrendByStore',this.ExpenseTrendByStore)

                if (this.ExpenseTrendByStoreKeys.length == 0) {
                  this.NoData = true;
                } else {
                  this.NoData = false;
                }
                this.shared.spinner.hide();
              } else {
                // this.toast.error('Empty Response','');
                this.shared.spinner.hide();
                this.NoData = true;
              }
            } else {
              // this.toast.error('Empty Response','');
              // //console.log( this.XpenseTrendByStoreKeys,this.ExpenseTrendByStoreKeys)

              this.ExpenseTrendByStore = [];
              this.AllDatakeys = [];
              this.shared.spinner.hide();
              this.NoData = true;
            }
          } else {
            let resTime = (new Date().getTime() - startFrom) / 1000;
            // this.toast.error(x.status, '');
            this.shared.spinner.hide();
            this.NoData = true;
            // this.logSaving(
            //   this.solutionurl + 'Price/GetIncomeStatementStoreComposite',
            //   obj,
            //   resTime,
            //   'Error'
            // );
          }
        },
        (error: any) => {
          // this.toast.error('502 Bad Gate Way Error', '');
          this.shared.spinner.hide();
          this.NoData = true;
        }
      );

    this.shared.api.GetHeaderData().subscribe((res: any) => {
      // //console.log('Filter Obj after Update : ', res);
      //this.componentDetails = res.obj;
    });
  }

  Fsdetails: any

  openDetails(Object: any, storename: any, ref: any, item: any) {

    // console.log(Object,storename);
    // this.stores =this.comm.groupsandstores.filter((v:any)=> v.sg_id == this.groups)[0].Stores
    // this.index = '';
    // // //console.log(Object, storename, ref, item, this.selectedstorename);
    // if (ref == 'store') {
    //   let index = this.stores.filter((store:any) => store.storename == storename);
    //   // //console.log(Object,storename,ref,index);
    //   this.selectedstorevalues = index[0].ID;
    //   this.selectedstorename = index[0].storename;


    // } else {
    //   if (this.selectedstorename == 'ALL STORES') {
    //     this.selectedstorevalues = 0;
    //   } else {
    //     let index = this.stores.filter(
    //       (store:any) => store.storename == this.selectedstorename
    //     );
    //     // //console.log(Object,storename,ref,index);

    //     this.selectedstorevalues = index[0].ID;
    //     this.selectedstorename = index[0].storename;
    //   }
    // const DetailsSF = this.ngbmodal.open(StoreDetailsComponent, {
    //   size: 'xl',
    //   backdrop: 'static',
    // });
    this.Fsdetails = {
      TYPE: Object.LABLEVAL,
      NAME: Object.LABLE,
      STORES: '2',
      LatestDate: this.Month,
      // Group: this.groups,
    };
    this.GetDetails()
  }






  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    }
    return false;
  }
  SFstate: any;
  ngAfterViewInit(): void {

    this.shared.api.getExportToExcelAllReports().subscribe((res: { obj: { state: boolean; title: string; }; }) => {
      this.SFstate = res.obj.state;
      if (res.obj.title == 'Income Statement Store Composite') {
        if (res.obj.state == true) {
          this.exportToExcel();
        }
      }
    });


    this.shared.api.getExportToPDFAllReports().subscribe((res: { obj: { title: string; statePDF: boolean; }; }) => {
      if (res.obj.title == 'Income Statement Store Composite') {
        if (res.obj.statePDF == true) {
          // this.generatePDF();
        }
      }
    });
    this.shared.api.getExportToPrintAllReports().subscribe((res: { obj: { title: string; statePrint: boolean; }; }) => {
      if (res.obj.title == 'Income Statement Store Composite') {
        if (res.obj.statePrint == true) {
          // this.GetPrintData();
        }
      }
    });
    this.shared.api.getExportToEmailPDFAllReports().subscribe((res: { obj: { title: string; stateEmailPdf: boolean; Email: any; notes: any; from: any; }; }) => {
      if (res.obj.title == 'Income Statement Store Composite') {
        if (res.obj.stateEmailPdf == true) {
          // this.sendEmailData(res.obj.Email, res.obj.notes, res.obj.from);
        }
      }
    });
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




  //-----------reports----------//
  minDate!: Date;
  maxDate!: Date;

  onOpenCalendar(container: any) {
    container.setViewMode('Month');
    container.monthSelectHandler = (event: any): void => {
      container.value = event.date;
      this.Month = event.date;
      return;
    };
  }
  changeDate(e: any) {
    ////console.log(e);
    this.Month = e;
  }


  activePopover: number = -1;
  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent) {
    const clickedInside = (event.target as HTMLElement).closest('.btn-secondary, .canvendar-select-popover');
    if (!clickedInside) {
      this.activePopover = -1;
    }
  }
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


    let date =
      ('0' + (this.Month.getMonth() + 1)).slice(-2) +
      '-01' +
      '-' +
      this.Month.getFullYear();


    ////console.log(date,this.month);
    let sname: any = 'All Stores';
    // if (this.selectedStore.toString() !== '0') {
    //   sname = this.stores.filter((e:any) => e.ID == this.selectedStore);
    //   sname = sname[0].DEALER_NAME;
    // }

    const data = {
      Reference: 'Income Statement Store Composite',


      storeValues: '2',
      month: date,
      Sname: sname,


    };

    //////console.log(data);
    this.shared.api.SetReports({
      obj: data,
    });
    this.GetData()

  }


  //------------Details-----------------//
  spinnerLoader: boolean = true;

  expandedIndex: number | null = null;
  FSSubDetailsMap: { [index: number]: any[] } = {};

  LatestDate: any;
  SelectMonthYear: any;
  filteredFSdetailsData: any[] = [];
  itemsPerPage: number = 100;
  currentPage: number = 1;
  searchText: string = '';
  FSDetailsData: any = [];
  Opacity: any = 'N';
  clickedPage: number | null = null;
  getPostingSubAmountTotal(index: number): number {
    const subRows = this.FSSubDetailsMap[index];
    if (!subRows || !Array.isArray(subRows)) {
      return 0;
    }

    return subRows.reduce((total, item) => {
      return total + (item.PostingAmount || 0);
    }, 0);
  }

  GetSubDetails(AcctNo: any, StoreName: any, index: number) {
    // Toggle: If same row clicked again, collapse it
    if (this.expandedIndex === index) {
      this.expandedIndex = null;
      return;
    }

    this.expandedIndex = index;

    this.spinnerLoader = true;
    // const format = 'MM-yyyy';
    // const locale = 'en-US';
    // const myDate = this.Fsdetails.LatestDate.replace(/-/g, '/');
    // const formattedDate = formatDate(myDate, format, locale);
    // this.LatestDate = formattedDate;

    const SelectedMY = this.shared.datePipe.transform(
      new Date(this.Fsdetails.LatestDate),
      'MMM-yy'
    );
    this.SelectMonthYear = SelectedMY;

    const Obj = {
      as_ids: StoreName,
      dept: '',
      subtype: '',
      subtypedetail: this.Fsdetails.NAME,
      FinSummary: this.Fsdetails.TYPE,
      date: this.shared.datePipe.transform(this.Month, 'MM-dd-yyyy'),
      accountnumber: AcctNo,
    };

    this.shared.api
      .postmethod(this.comm.routeEndpoint + 'GetFinancialSummaryDetails', Obj)
      .subscribe((res: { status: number; response: any[]; }) => {
        this.spinnerLoader = false;
        if (res.status === 200) {
          this.FSSubDetailsMap[index] = res.response;
        }
      });
  }
  get postingAmountTotal(): number {
    return this.filteredFSdetailsData.reduce((total: any, item: { PostingAmount: any; }) => {
      return total + (item.PostingAmount || 0);
    }, 0);
  }
  get paginatedItems() {
    const startIndex = (this.currentPage - 1) * this.itemsPerPage;
    const endIndex = startIndex + this.itemsPerPage;
    return this.filteredFSdetailsData.slice(startIndex, endIndex);
  }
  filterData() {
    if (this.searchText.trim() !== '') {
      this.filteredFSdetailsData = this.FSDetailsData.filter(
        (item: any) =>
          (item.Store &&
            item.Store.toLowerCase().includes(this.searchText.toLowerCase())) ||
          (item.accountnumber &&
            item.accountnumber
              .toLowerCase()
              .includes(this.searchText.toLowerCase())) ||
          (item.Control &&
            item.Control
              .toLowerCase()
              .includes(this.searchText.toLowerCase())) ||
          (item.accountingdate &&
            item.accountingdate
              .toLowerCase()
              .includes(this.searchText.toLowerCase())) ||
          (item.Make &&
            item.Make.toLowerCase().includes(this.searchText.toLowerCase())) ||
          (item.detaildescription &&
            item.detaildescription
              .toLowerCase()
              .includes(this.searchText.toLowerCase()))
      );
    } else {
      this.filteredFSdetailsData = [...this.FSDetailsData];
    }
    this.currentPage = 1;
  }
  GetDetails() {
    this.NoData = false;
    this.Opacity = 'N';
    const format = 'MM-yyyy';
    const locale = 'en-US';
    // const myDate = this.Fsdetails.LatestDate.replace(/-/g, '/');
    // const formattedDate = formatDate(myDate, format, locale);
    // this.LatestDate = formattedDate;
    const SelectedMY = this.shared.datePipe.transform(
      new Date(this.Fsdetails.LatestDate),
      'MMM-yy'
    );
    this.SelectMonthYear = SelectedMY;
    // const Obj = {
    //   Type: this.Fsdetails.TYPE,
    //   Date: this.LatestDate,
    //   Stores:
    //     this.Fsdetails.STORES == undefined ||
    //     this.Fsdetails.STORES == null ||
    //     this.Fsdetails.STORES == 0
    //       ? ''
    //       : this.Fsdetails.STORES,
    //   PageNumber: this.pageNumber,
    //   PageSize: '100',
    // };
    const Obj = {
      "as_ids": '2',
      "dept": "",
      "subtype": "",
      "subtypedetail": this.Fsdetails.NAME,
      "FinSummary": this.Fsdetails.TYPE,
      "date": this.shared.datePipe.transform(this.Month, 'MM-dd-yyyy'),
      "accountnumber": ""
    }
    console.log(Obj);

    this.shared.api
      .postmethod(this.comm.routeEndpoint + 'GetFinancialSummaryDetails', Obj)
      .subscribe((res: { status: number; response: any; }) => {
        if (res.status == 200) {
          this.FSDetailsData = res.response;
          this.filterData();
          console.log(this.FSDetailsData);
          this.spinnerLoader = false;
          // if (this.FSDetailsData.length > 0) {
          //   this.NoData = false;
          // } else {
          this.NoData = true;
          // }
        }
      });
  }
  getMaxPageNumber(): number {
    return Math.ceil(this.filteredFSdetailsData.length / this.itemsPerPage);
  }
  goToLastPage() {
    this.currentPage = this.getMaxPageNumber();
    this.clickedPage = null;
    this.expandedIndex = null;
  }
  nextPage() {
    if (this.currentPage < this.getMaxPageNumber()) {
      this.currentPage++;
      this.clickedPage = null;
      this.expandedIndex = null;
    }
  }
  getStartRecordIndex(): number {
    return (this.currentPage - 1) * this.itemsPerPage;
  }
  getEndRecordIndex(): number {
    const endIndex = this.getStartRecordIndex() + this.itemsPerPage;
    return endIndex > this.FSDetailsData.length
      ? this.FSDetailsData.length
      : endIndex;
  }
  prevPage() {
    if (this.currentPage > 1) {
      this.currentPage--;
      this.clickedPage = null;
      this.expandedIndex = null;
    }
  }
  goToFirstPage() {
    this.currentPage = 1;
    this.clickedPage = null;
    this.expandedIndex = null;
  }
  close() {
    // this.ngbmodalActive.close();
    // console.log(this.Opacity);
    this.filteredFSdetailsData = [];
    if (this.Fsdetails.STORES != '') {
      this.goToFirstPage();
    }
  }
  exportToExcel(): void {
    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('Income Statement Store Composite');


    const title = worksheet.addRow(['Income Statement Store Composite Report']);
    title.font = { size: 14, bold: true, name: 'Arial' };
    title.alignment = { vertical: 'middle', horizontal: 'center' };
    worksheet.mergeCells('A1:D1');
    worksheet.addRow([]);


    const formattedMonth = this.shared.datePipe.transform(this.Month, 'MMMM-yyyy');
    const filters = [
      { name: 'Store:', values: 'WesternAuto' },
      { name: 'Month:', values: formattedMonth },
    ];

    let startIndex = 3;
    filters.forEach((filter) => {
      startIndex++;
      worksheet.addRow('');
      worksheet.getCell(`A${startIndex}`).value = filter.name;
      worksheet.getCell(`B${startIndex}`).value = filter.values;
    });

    worksheet.addRow([]);


    const formattedMonthHeader = this.shared.datePipe.transform(this.Month, 'MMMM-yy');
    const headerLabels = [formattedMonthHeader, ...this.ExpenseTrendByStoreKeys.slice(0, -1)].map(label => label.toString().toUpperCase())

    const headerRow = worksheet.addRow(headerLabels);
    headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
    // headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2F5597' } };
    headerRow.height = 22;


    headerRow.eachCell((cell, colNumber) => {
      if (colNumber <= headerLabels.length) { // Only style cells with header text
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FF2F5597' } // Blue background
        };
      }
    });



    this.ExpenseTrendByStore.forEach((item: any) => {
      const rowData: any[] = [];


      rowData.push(item.LABLE || '-');


      this.AllDatakeys.slice(0, -1).forEach((key: string) => {
        let value = item[key];

        if (value == null || value === 0) {
          value = '-';
        } else if (item.LABLE.endsWith('%')) {

          value = value < 0
            ? `(${Math.abs(value).toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 0 })}%)`
            : `${value.toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 0 })}%`;
        } else if (['New Units', 'Pre-Owned Units', 'Unit Retail Sales', 'Fleet Units'].includes(item.LABLE)) {

          value = value.toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 0 });
        } else if (typeof value === 'number') {

          value = value < 0
            ? `($${Math.abs(value).toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 0 })})`
            : `$${value.toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 0 })}`;
        }

        rowData.push(value);
      });

      worksheet.addRow(rowData);
    });


    worksheet.columns.forEach((col: any) => {
      col.width = 20;
    });


    workbook.xlsx.writeBuffer().then((buffer: any) => {
      const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      this.shared.exportToExcel(workbook, 'Income Statement Store Composite');
    });
  }

  onExcelClick() {
    const FSDetailsData = [...this.filteredFSdetailsData];
    const FSSubDetailsMap = this.FSSubDetailsMap;

    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('Income Statement Store Composite Details');


    const titleRow = worksheet.addRow(['Income Statement Store Composite Details']);
    titleRow.font = { name: 'Arial', size: 14, bold: true };
    titleRow.alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.mergeCells('A1:N1');
    worksheet.addRow([]);

    const filters = [
      { name: 'Selected Details::', values: '' },
      { name: 'Type:', values: this.Fsdetails.NAME },
      { name: 'Date:', values: this.LatestDate },
      { name: 'Store:', values: 'WesternAuto' },
    ];


    let startIndex = 3;
    filters.forEach(f => {
      startIndex++;
      worksheet.addRow('');
      worksheet.getCell(`A${startIndex}`).value = f.name;
      worksheet.getCell(`B${startIndex}`).value = f.values;
      worksheet.getCell(`A${startIndex}`).font = { bold: true };
      worksheet.mergeCells(`B${startIndex}:C${startIndex}`);
    });
    worksheet.addRow([]);


    const headerRow1 = [
      'Store Name',
      'Account Number',
      'Description',
      `Balance (${this.SelectMonthYear})`,
    ];
    const headerRow = worksheet.addRow(headerRow1);
    headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
    headerRow.height = 22;
    headerRow.eachCell((cell, colNumber) => {
      if (colNumber <= headerRow1.length) { // Only style cells with header text
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FF2F5597' } // Blue background
        };
      }
    });




    // Setup Excel
    const DATE_EXTENSION = this.shared.datePipe.transform(new Date(), 'MMddyyyy');
    const DateToday = this.shared.datePipe.transform(new Date(), 'MM/dd/yyyy h:mm:ss a');
    worksheet.views = [{ state: 'frozen', ySplit: 9, topLeftCell: 'A10', showGridLines: false }];



    let totalBalance = 0;
    const columnWidths = new Array(headerRow1.length).fill(10);

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
            fgColor: { argb: 'FF2F5597' },
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

    // Export to Excel File
    // workbook.xlsx.writeBuffer().then((data) => {
    //   const blob = new Blob([data], {
    //     type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    //   });
    //   FileSaver.saveAs(blob, `Financial_Summary_Details_${DATE_EXTENSION}.xlsx`);
    // });

      workbook.xlsx.writeBuffer().then((buffer: any) => {
      const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      this.shared.exportToExcel(workbook, `Income Statement Store Composite_Details_${DATE_EXTENSION}.xlsx`);
    });
  }
}
