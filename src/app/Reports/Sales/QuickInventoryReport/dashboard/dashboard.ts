import { Component, OnInit, ViewChild, ElementRef, HostListener } from '@angular/core';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { Setdates } from '../../../../Core/Providers/SetDates/setdates';
import { BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { common } from '../../../../common';
import { DateRangePicker } from '../../../../CommonFilters/date-range-picker/date-range-picker';
import { Subscription } from 'rxjs';
import { FilterPipe } from '../../../../Core/Providers/filterpipe/filter.pipe';
import { Options, LabelType } from "@angular-slider/ngx-slider";
import { NgxSliderModule } from '@angular-slider/ngx-slider';
@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [SharedModule, BsDatepickerModule,NgxSliderModule],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss'
})
export class Dashboard {
  reportOpenSub!: Subscription;
  reportGetting!: Subscription;
  Pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;
  date: any;
  StoreValues: any = [];
  NoData: boolean = false;
  Month: any = '';
  StoreName: any = 'All Stores';
  DealType: any = ['N', 'U', 'C'];
  AgeFrom: any = 0;
  AgeTo: any = 300;
  QISearchName: any;
  PresentDayDate: any;
  groups: any = 1;
  header: any = [{ type: 'Bar', StoreValues: this.StoreValues, Month: this.Month, DealType: this.DealType, AgeFrom: this.AgeFrom, AgeTo: this.AgeTo, groups: this.groups }];
  popup: any = [{ type: 'Popup' }];
  @ViewChild('content', { static: false }) content!: ElementRef;
  constructor(
    public shared: Sharedservice, public setdates: Setdates, private comm: common
  ) {

    
    this.shared.setTitle(this.shared.common.titleName + '-Inventory Browser');
    this.date = new Date();


    if (localStorage.getItem('Fav') != 'Y') {
      const data = {
        title: 'Inventory Browser',
        path1: '',
        path2: '',
        path3: '',
        Month: this.date,
        stores: this.StoreValues.toString(),
        store: this.StoreValues.toString(),
        AgeFrom: this.AgeFrom,
        AgeTo: this.AgeTo,
        DealType: this.DealType,
        groups: this.groups,
        completestores: this.completeData,
        count: 0
      };
      this.shared.api.SetHeaderData({ obj: data });
      this.header = [{ type: 'Bar', StoreValues: this.StoreValues, Month: this.date, DealType: this.DealType, AgeFrom: this.AgeFrom, AgeTo: this.AgeTo, groups: this.groups, completestores: this.completeData, }]
      this.GetQuickInventoryData();
    } else {
      this.getFavReports();
    }
  }




  ngOnInit(): void {

    // var curl = 'https://fbxtract.axelautomotive.com/favouritereports/GetQuickInventory';
    // this.apiSrvc.logSaving(curl, {}, '', 'Success', 'Inventory Browser');
  }
  completeData: any = []

  Favreports: any = [];

  getFavReports() {
    const obj = {
      Id: localStorage.getItem('Fav_id'),
      expression: '',
    };
    this.shared.api.postmethod('favouritereports/get', obj).subscribe((res) => {
      if (res.status == 200) {
        if (res.response.length > 0) {
          this.Favreports = res.response;

          let age = this.Favreports[2].Fav_Report_Value.split(',');
          this.AgeFrom = age[0];
          this.AgeTo = age[1];
          this.StoreValues = this.Favreports[0].Fav_Report_Value;
          this.DealType = this.Favreports[1].Fav_Report_Value;
          this.GetQuickInventoryData();
          localStorage.setItem('Fav', 'N');
          const data = {
            title: 'Inventory Browser',
            path1: '',
            path2: '',
            path3: '',
            Month: this.date,
            stores: this.StoreValues,
            // store: this.StoreValues,
            AgeFrom: this.AgeFrom,
            AgeTo: this.AgeTo,
            DealType: this.DealType,
          };
          // console.log(data);
          this.shared.api.SetHeaderData({ obj: data });
        }
      }
    });
  }
  QuickInventoryData: any = [];
  NodataFound: boolean = false;
  currentPage: number = 1;
  CurrentPageSetting: number = 1
  itemsPerPage: number = 100;
  maxPageButtonsToShow: number = 3;
  clickedPage: number | null = null;
  filteredQuickInventoryData: any[] = [];
  searchText: string = '';
  GetQuickInventoryData() {
    this.shared.spinner.show();
    let obj = {};
    // console.log('locstg', JSON.parse(localStorage.getItem('IBObject')!));
    if (localStorage.getItem('IBObject') != null) {
      // console.log('locstg', JSON.parse(localStorage.getItem('IBObject')!));
      const InvObj = JSON.parse(localStorage.getItem('IBObject')!);
      obj = {
        DealerId: InvObj.invobj.store,
        DealType: InvObj.invobj.DealType,
        FromAge: InvObj.invobj.FromAge,
        ToAge: InvObj.invobj.ToAge,
        UserID: 0,
      };
    } else {
      obj = {
        DealerId: '2',
        // DealerId: '8',

        DealType: this.DealType.join(','),
        FromAge: this.AgeFrom,
        ToAge: this.AgeTo,
        UserID: 0,
      };
    }

    // const obj = {
    //   "DealerId": '',
    // }
    console.log(obj);
    const curl = this.shared.getEnviUrl() + this.comm.routeEndpoint +  'GetCITFloorPlanData';
    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetQuickInventory', obj).subscribe(
      (res) => {

        const currentTitle = document.title;
        this.shared.api.logSaving(curl, {}, '', res.message, currentTitle);
        if (res.status == 200) {
          localStorage.removeItem('IBObject');
          this.QuickInventoryData = res.response.map((v: any) => ({
            ...v, notesView: '+'
          }));
          this.QuickInventoryData.some(function (x: any) {
            if (x.Notes != undefined && x.Notes != null && x.Notes != '') {
              x.Notes = JSON.parse(x.Notes);
            }
            if (x.Comments != undefined && x.Comments != null && x.Comments != '') {
              x.Comments = JSON.parse(x.Comments);
            }
          });
          this.filterData();
          console.log('QI Data', this.QuickInventoryData);
          this.shared.spinner.hide();
          this.NodataFound = true;
          if (this.QuickInventoryData.length > 0) {
            this.NoData = false;
          } else {
            this.NoData = true;
          }
        } else {
          this.shared.spinner.hide();
          // this.toast.error('Invalid Details');
        }
      },
      // (error) => {
      //   console.log(error);
      // }
    );
  }
  toggleView(data: any) {
    if (data.notesView == '+') {
      data.notesView = '-'
    } else {
      data.notesView = '+'
    }
  }
  notesViewState: boolean = true
  notesView() {
    this.notesViewState = !this.notesViewState
  }
  notesData: any = {}
  Notespopup: any
  addNotes(data: any, ref: any) {
    this.scrollpositionstoring = this.scrollCurrentposition

    this.notesData = {
      // store: data.StoreID,
      // mainkey: data.StockNo,
      // module: 'IN',
      // apiRoute:'AddNotesAction'
      store: '2',
      title1: data.StockNo,
      title2: '',
      apiRoute: 'AddGeneralNotes'
    }
    this.Notespopup = this.shared.ngbmodal.open(ref, { size: 'xxl', backdrop: 'static' });
  }
  
  callLoadingState = 'FL'
  closeNotes(e: any) {
    // this.ngbmodalActive.dismiss()
    this.Notespopup.close()
    if (e == 'S') {
      console.log(this.currentPage);
      this.CurrentPageSetting = this.currentPage
      this.callLoadingState = 'ANS'
      this.GetQuickInventoryData();
    }
  }

  // downloadPDF() {
  //   this.generatePDF()
  //   // this.apiSrvc.generatePdf(this.content.nativeElement, 'my-pdf')
  // }
  filterData() {
    if (this.searchText.trim() !== '') {
      this.filteredQuickInventoryData = this.QuickInventoryData.filter((item: any) =>
        (item.Store && item.Store.toLowerCase().includes(this.searchText.toLowerCase())) ||
        (item.StockNo && item.StockNo.toLowerCase().includes(this.searchText.toLowerCase())) ||
        (item.DealType && item.DealType.toLowerCase().includes(this.searchText.toLowerCase())) ||
        (item.Make && item.Make.toLowerCase().includes(this.searchText.toLowerCase())) ||
        (item.Model && item.Model.toLowerCase().includes(this.searchText.toLowerCase())) ||
        (typeof item.Age === 'string' && item.Age.toLowerCase().includes(this.searchText.toLowerCase())) ||
        (item.VIN && item.VIN.toLowerCase().includes(this.searchText.toLowerCase())) ||
        (typeof item.Miles === 'string' && item.Miles.toLowerCase().includes(this.searchText.toLowerCase()))
      );
    } else {
      this.filteredQuickInventoryData = [...this.QuickInventoryData];
    }
    this.callLoadingState == 'ANS' ? this.sort(this.column, this.filteredQuickInventoryData, this.callLoadingState) : ''
    let position = this.scrollpositionstoring + 10
    setTimeout(() => {
      this.scrollcent.nativeElement.scrollTop = position
    }, 500);
  }

  clearSearchFilter() {
    this.searchText = '';
    this.filterData();
  }

  get paginatedItems() {
    this.CurrentPageSetting != 1 ? this.currentPage = this.CurrentPageSetting : '';
    const startIndex = (this.currentPage - 1) * this.itemsPerPage;
    const endIndex = startIndex + this.itemsPerPage;
    return this.filteredQuickInventoryData.slice(startIndex, endIndex);
  }

  getMaxPageNumber(): number {
    return Math.ceil(this.filteredQuickInventoryData.length / this.itemsPerPage);
  }

  nextPage() {
    if (this.currentPage < this.getMaxPageNumber()) {
      this.currentPage++;
      this.clickedPage = null;
    }
    this.CurrentPageSetting = 1;

  }

  prevPage() {
    if (this.currentPage > 1) {
      this.currentPage--;
      this.clickedPage = null;
    }
    this.CurrentPageSetting = 1;

  }

  goToFirstPage() {
    this.currentPage = 1;
    this.clickedPage = null;
    this.CurrentPageSetting = 1;
  }

  getPageNumbers(): number[] {
    const maxPage = this.getMaxPageNumber();
    const currentPage = this.currentPage;
    const maxButtonsToShow = this.maxPageButtonsToShow;
    if (maxPage <= 0) return [];

    let start = Math.max(currentPage - Math.floor(maxButtonsToShow / 2), 1);
    let end = start + maxButtonsToShow - 1;
    if (end > maxPage) {
      end = maxPage;
      start = Math.max(end - maxButtonsToShow + 1, 1);
    }
    return Array.from({ length: end - start + 1 }, (_, index) => start + index);
  }


  goToPage(page: number) {
    this.currentPage = page;
    this.clickedPage = null;
    this.CurrentPageSetting = 1;

  }
  goToLastPage() {
    this.currentPage = this.getMaxPageNumber();
    this.clickedPage = null;
    this.CurrentPageSetting = 1;

  }
  getStartRecordIndex(): number {
    return (this.currentPage - 1) * this.itemsPerPage;
  }
  getEndRecordIndex(): number {
    const endIndex = this.getStartRecordIndex() + this.itemsPerPage;
    return endIndex > this.QuickInventoryData.length ? this.QuickInventoryData.length : endIndex;
  }
  QIClick: any;
  ngAfterViewInit(): void {
    this.reportOpenSub = this.shared.api.GetReportOpening().subscribe((res) => {
      if (this.reportOpenSub != undefined) {
        if (res.obj.Module == 'Inventory Browser') {
          document.getElementById('report')?.click();
        }
      }
    });
    this.reportGetting = this.shared.api.GetReports().subscribe((data) => {
      if (this.reportGetting != undefined) {

        if (data.obj.Reference == 'Inventory Browser') {
          if (data.obj.header == undefined) {
            this.date = data.obj.month;
            this.Month = data.obj.month;
            this.StoreValues = data.obj.storeValues;
            this.StoreName = data.obj.Sname;
            this.AgeFrom = data.obj.AgeFrom;
            this.AgeTo = data.obj.AgeTo;
            this.DealType = data.obj.DealType;
            this.groups = data.obj.groups;
          } else {
            if (data.obj.header == 'Yes') {
              this.StoreValues = data.obj.storeValues;
            }
          }
          if (this.StoreValues != '') {
            this.GetQuickInventoryData();
            this.goToFirstPage();
          } else {
            this.NoData = true
            this.filteredQuickInventoryData = []
          }

          const headerdata = {
            title: 'Inventory Browser',
            path1: '',
            path2: '',
            path3: '',
            Month: this.Month,
            stores: this.StoreValues,
            AgeFrom: this.AgeFrom,
            AgeTo: this.AgeTo,
            DealType: this.DealType,
            groups: this.groups,
            completestores: this.completeData,
          };
          // // console.log(headerdata)
          this.shared.api.SetHeaderData({
            obj: headerdata,
          });
          this.header = [{ type: 'Bar', StoreValues: this.StoreValues, Month: new Date(this.Month), DealType: this.DealType, AgeFrom: this.AgeFrom, AgeTo: this.AgeTo, groups: this.groups, completestores: this.completeData, }]
        }
      }
    });
    this.excel = this.shared.api.getExportToExcelAllReports().subscribe((res) => {
      if (this.excel != undefined) {
        this.QIClick = res.obj.state;
        if (res.obj.title == 'Inventory Browser') {
          if (res.obj.state == true) {
            this.ExportToExcelQI();
          }
        }
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res) => {
      if (this.print != undefined) {
        if (res.obj.title == 'Inventory Browser') {
          if (res.obj.statePrint == true) {
            // this.GetPrintData();
          }
        }
      }
    });
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res) => {
      if (this.email != undefined) {
        if (res.obj.title == 'Inventory Browser') {
          if (res.obj.stateEmailPdf == true) {
            // this.sendEmailData(res.obj.Email, res.obj.notes, res.obj.from);
          }
        }
      }
    });
    this.Pdf = this.shared.api.getExportToPDFAllReports().subscribe((res) => {
      if (this.Pdf != undefined) {
        if (res.obj.title == 'Inventory Browser') {
          if (res.obj.statePDF == true) {
            // this.generatePDF();
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
  reportOpen(temp: any) {
    // this.ngbmodalActive = this.ngbmodal.open(temp, {
    //   size: 'xl',
    //   backdrop: 'static',
    // });
  }


  Scrollpercent: any = 0;
  scrollCurrentposition: any = 0;
  scrollpositionstoring: any = 0;
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

  close(data: any) {
    // console.log(data);
    this.index = '';
  }

  selBlock: any;
  screenheight: any = 0; divheight: any = 0; trposition: any = 0;
  commentopen(item: any, i: any, slblock: any = '') {
    this.screenheight = window.screen.height;
    this.divheight = (<HTMLInputElement>document.getElementById('scrollcent')).offsetHeight;
    this.trposition = (<HTMLInputElement>document.getElementById(i)).offsetTop;
    this.index = '';
    console.log('Selected Obj :', item);
    //return
    this.selBlock = slblock + i.toString();
    this.index = i.toString();
    this.commentobj = {
      TYPE: item.StockNo,
      NAME: item.StockNo,
      STORES: i,
      STORENAME: item.StockNo,
      Month: '',
      ModuleId: '79',
      ModuleRef: 'IB',
      state: 1,
      indexval: i,
      apiRoute: 'AddNotesAction'
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

    //     // // on dismiss
    //     // const Data = {
    //     //   state: true,
    //     // };
    //     // this.apiSrvc.setBackgroundstate({ obj: Data });
    //     this.GetData();
    //   }
    // );

  }

  index = '';
  commentobj: any = {};

  commentsVisibility: boolean = true
  openComments() {
    this.commentsVisibility = !this.commentsVisibility
  }
  addcmt(data: any) {


    if (data == 'AD') {
      this.GetQuickInventoryData();
    }
    // if (data == 'AD') {
    //   if (this.Filter == 'StoreSummary') {
    //     this.GetQuickInventoryData();
    //   }
    //   if (this.Filter == 'StoreSummary') {
    //     this.GetQuickInventoryData();
    //   }
    // }
  }


  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    }
    return false;
  }

  isDesc: boolean = false;
  column: string = 'CategoryName';

  sort(property: any, data: any, state?: any) {
    if (state == undefined) {
      this.isDesc = !this.isDesc;
    }
    this.callLoadingState = 'FL'
    this.column = property;
    let direction = this.isDesc ? 1 : -1;
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

//----------Reports-------------//
activePopover: number = -1;
Performance: any = 'Load';
tab = 'S';
@HostListener('document:click', ['$event'])
onDocumentClick(event: MouseEvent) {
  const clickedInside = (event.target as HTMLElement).closest('.dropdown-toggle, .reportstores-card, .timeframe');
  if (!clickedInside) {
    this.activePopover = -1;
  }
}
togglePopover(popoverIndex: number) {
  if (this.activePopover === popoverIndex) {
    this.activePopover = -1;
  } else {
    this.activePopover = popoverIndex;
  }
}
getDealTypeLabel(): string {
  if (this.DealType.length === 3) {
    return 'All Types';
  } else if (this.DealType.length === 1) {
    return this.getDealTypeName(this.DealType[0]);
  } else {
    return 'Selected';
  }
}

multipleorsingle(e: any) {
  const index = this.DealType.findIndex((i: any) => i == e);
  if (index >= 0) {
    this.DealType.splice(index, 1);
  } else {
    this.DealType.push(e);
  }
  // console.log(this.DealType);
}
// AgeFrom: any = '';
// AgeTo: any = '';

value: number = 0;
highValue: number = 1000;
optionProx: Options = {
  floor: 0, ceil: 1000,
  showSelectionBarFromValue: 0,
  hideLimitLabels: true,
  translate: (value: number, label: LabelType): string => {
    switch (label) {
      case LabelType.Low: return "<a>" + value + " Age</a>";
      case LabelType.High:
        return " " + value + " Age";
      default:
        return "" + value + " Age";
    }
  }
};
getDealTypeSpan(): string {
  if (this.DealType.length === 3) {
    return '';
  } else if (this.DealType.length === 1) {
    return '';
  } else {
    return `(${this.DealType.length})`;
  }
}

getDealTypeName(dealType: string): string {
  switch (dealType) {
    case 'N':
      return 'New';
    case 'U':
      return 'Used';
    case 'C':
      return 'CPO';
    default:
      return '';
  }
}
Tabs(e: any) {
  this.tab = e;
}
FT_Name: any;
SFT_Name: any;
selectedStoresRegions: any = [];
indexstores: any = [];

viewreport() {
  this.activePopover = -1
 
    const data = {
      Reference: 'Inventory Browser',
      storeValues: '2',
      month: '',
      storeName: '',
      AgeFrom: this.AgeFrom,
      AgeTo: this.AgeTo,
      DealType: this.DealType,
      // groups: this.selectedGroups
    };
    // console.log(data);
    this.shared.api.SetReports({
      obj: data,
    });
    this.Performance = 'Unload';
    // this.close();
  }
  ExportToExcelQI(): void {
    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('Inventory Browser');
  
  
    const dealTypeMaps: Record<string, string> = {
      N: 'New',
      U: 'Used',
      C: 'Certified Pre-Owned'
    };
    
    // 🔹 Convert DealType (which could be string, array, or single value) safely
    let dealTypesFormatted = '-';
    
    if (this.DealType) {
      if (Array.isArray(this.DealType)) {
        // Case: DealType = ['N', 'U', 'C']
        dealTypesFormatted = this.DealType
          .map((t: string) => dealTypeMaps[t.trim()] || t)
          .join(', ');
      } else if (typeof this.DealType === 'string') {
        // Case: DealType = 'N,U,C' or 'N'
        dealTypesFormatted = this.DealType
          .split(',')
          .map((t: string) => dealTypeMaps[t.trim()] || t)
          .join(', ');
      } else {
        // Case: Single non-string value like N
        const key = String(this.DealType).trim();
        dealTypesFormatted = dealTypeMaps[key] || key;
      }
    }
    
    const filters: any = [
      { name: 'Store :', values: 'Silvertip' },
      { name: 'Stock Type:', values: dealTypesFormatted }
    ];
    
  
    const ReportFilter = worksheet.addRow(['Report Controls :']);
    ReportFilter.font = { name: 'Arial', size: 10, bold: true };
  
    let startIndex = 2;
    filters.forEach((val: any) => {
      startIndex++;
      worksheet.addRow('');
      worksheet.mergeCells(`B${startIndex}:C${startIndex}`);
      worksheet.getCell(`A${startIndex}`).value = val.name;
      worksheet.getCell(`B${startIndex}`).value = val.values;
    });
    worksheet.addRow('');
  
 
    const secondHeader = [
      'Store', 'StockNo', 'Stock Type', 'Year', 'Make', 'Model',
      'Color', 'Interior Color', 'Age', 'List Price', 'MSRP', 'Miles', 'VIN'
    ];
    const headerRow = worksheet.addRow(secondHeader);
    headerRow.height = 25;
    headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
    headerRow.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF2F5597' }
    };
  
   
    const bindingHeaders = [
      'Store', 'StockNo', 'DealType', 'Year', 'Make', 'Model',
      'Color', 'Interior Color', 'Age', 'ListPrice', 'Internet', 'Miles', 'VIN'
    ];
  
  
    const dealTypeMap: Record<string, string> = {
      N: 'New',
      U: 'Used',
      C: 'Certified Pre-Owned'
    };
    const currencyFields = ['ListPrice', 'Internet'];
  
   
    for (const info of this.paginatedItems) {
      const rowData = bindingHeaders.map(key => {
        let val = info[key];
  
      
        if (key === 'DealType') {
          val = dealTypeMap[val] || '-';
        }
  
     
        if (val === 0 || val == null || val === '') {
          val = '-';
        }
  
        return val;
      });
  
      const dealerRow = worksheet.addRow(rowData);
  

      bindingHeaders.forEach((key, index) => {
        const cell = dealerRow.getCell(index + 1);
  
        if (currencyFields.includes(key)) {
       
          const num = Number(info[key]);
          if (!isNaN(num)) {
            cell.value = num;
            cell.numFmt = '"$"#,##0';
            cell.alignment = { horizontal: 'right' };
          } else {
            cell.value = '-';
          }
        } else if (!isNaN(Number(cell.value))) {
          cell.alignment = { horizontal: 'center' };
        } else {
          cell.alignment = { horizontal: 'left' };
        }
      });
    }
  
  
    worksheet.columns.forEach((column: any) => {
      let maxLength = 0;
      column.eachCell({ includeEmpty: true }, (cell: any) => {
        const cellValue = cell.value ? cell.value.toString() : '';
        maxLength = Math.max(maxLength, cellValue.length);
      });
      column.width = maxLength + 4;
    });
  

    workbook.xlsx.writeBuffer().then((buffer: any) => {
      const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      this.shared.exportToExcel(workbook, 'Inventory Browser');
    });
  }
  
}




  

