import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { Component, HostListener, } from '@angular/core';

import { Setdates } from '../../../../Core/Providers/SetDates/setdates';
import { CardealsReports } from '../cardeals-reports/cardeals-reports';
import { Subscription } from 'rxjs';
import { Notes } from '../../../../Layout/notes/notes';

@Component({
  selector: 'app-dashboard',
  imports: [SharedModule, CardealsReports, Notes],
  standalone: true,
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss'
})
export class Dashboard {
  Viewmore: boolean = false
  spinnerLoader: boolean = false
  notesViewState: boolean = false
  FromDate: any = '';
  ToDate: any = '';
  TotalReport: any = 'T';
  NoData!: boolean;
  CompleteComponentState: boolean = true;
  store: any = '1';
  QISearchName: any = '';
  path1: any = '';
  path2: any = '';
  path3: any = '';
  path1name: any = '';
  path2name: any = '';
  path3name: any = '';
  path1id: any;
  path2id: any;
  path3id: any;
  CurrentDate = new Date();
  groups: any = 0;
  GridView = 'Global';
  dealType: any = ['New', 'Used'];
  saleType: any = ['Retail', 'Lease', 'Misc', 'Special Order']
  dealStatus: any = ['Delivered', 'Capped', 'Finalized'];
  count = 0;
  storeName: any = ''
  DateType: any = 'MTD'

  header: any = [{
    type: 'Bar', storeIds: this.store, fromDate: this.FromDate, toDate: this.ToDate, groups: this.groups, datevaluetype: this.DateType
  }]
  pageNumber: any = 0;

  constructor(public shared: Sharedservice, public setdates: Setdates) {

    this.shared.setTitle(this.shared.common.titleName + '-Car Deals')

    // if (typeof window !== 'undefined') {
    // if (localStorage.getItem('UserDetails') != null) {
    if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
      this.groups = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Preferences
      this.store = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids.split(',')
    }

    this.setDates('MTD')
    this.DateType = 'MTD'
    // }
    // if (localStorage.getItem('Fav') != 'Y') {
    this.setHeaderReportData()
    // }
    // }
  }

  ngOnInit(): void {
    // localStorage.setItem('time', 'MTD');
    // this.reportOpenSub.unsubscribe()
    // this.reportGetting.unsubscribe()
    // this.Pdf.unsubscribe()
    // this.print.unsubscribe()
    // this.email.unsubscribe()
    // this.excel.unsubscribe()
  }
  details: any = [];
  otherblocks: any = ['Retail and Lease']
  obTemporary: any = 'Retail and Lease'
  callLoadingState = 'FL'
  cardealsdata: any = []
  getCardealsFilter() {
    this.details = []
    this.cardealsdata = []
    this.GetData()
  }

  setDates(type: any) {
    let dates: any = this.setdates.setDates(type)
    this.FromDate = dates[0];
    this.ToDate = dates[1];
    localStorage.setItem('time', type);
  }
  GetData() {
    this.shared.spinner.show();
    console.log(this.details, this.cardealsdata);
    this.spinnerLoader = true
    const obj = {
      startdealdate: this.FromDate,
      enddealdate: this.ToDate,
      Stores: this.store,
      dealtype: this.dealType.toString(),
      saletype: this.saleType.toString(),
      dealstatus: this.dealStatus.toString(),
      PageNumber: this.pageNumber,
      PageSize: '1000',
      ExtraType: this.otherblocks.toString(),
      SearchExp: this.QISearchName
    };
    this.count++
    const curl = this.shared.getEnviUrl() + this.shared.common.routeEndpoint + + 'GetSalesCarDeals';
    this.shared.api.postmethod(this.shared.common.routeEndpoint + 'GetSalesCarDeals', obj).subscribe(
      (res) => {
        const currentTitle = document.title;
        this.shared.api.logSaving(curl, {}, '', res.message, currentTitle);
        if (res.status == 200) {
          this.shared.spinner.hide()
          if (res.response != undefined) {
            if (res.response.length > 0) {
              this.details = res.response;
              this.details.some(function (x: any) {
                if (x.Notes != undefined && x.Notes != null) {
                  x.Notes = JSON.parse(x.Notes);
                }
              });
              this.cardealsdata = [
                ...this.cardealsdata,
                ...this.details,
              ];
              // this.callLoadingState == 'ANS' ? this.sort(this.column) : ''
              this.NoData = false;
              this.spinnerLoader = false
              console.log(this.details, this.cardealsdata);
            }
            else {
              // this.toast.error('Empty Response','');
              this.details = []
              this.shared.spinner.hide();
              if (this.cardealsdata.length == 0) {
                this.NoData = true;
              }
              this.spinnerLoader = false

            }
          } else {
            // this.toast.error('Empty Response');
            this.shared.spinner.hide();
            if (this.cardealsdata.length == 0) {
              this.NoData = true;
            }
            this.spinnerLoader = false
          }
        } else {
          this.spinnerLoader = false
          this.shared.spinner.hide();
          this.NoData = true;

        }
      },
      (error) => {

        this.shared.spinner.hide();

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

  // getTotal(frontgross: any, colname: any) {
  //   let total: any = 0
  //   frontgross.some(function (x: any) {
  //     total += parseInt(x[colname])
  //   })
  //   return total
  // }

  getTotal(frontgross: any[], colname: string) {
    return frontgross.reduce((total, x) => {
      const raw = x?.[colname];

      // Skip null/undefined/empty strings
      if (raw === null || raw === undefined || raw === '') return total;

      // Parse as number (handles strings like "123" or "123.45")
      const n = Number(raw);

      // Skip NaN
      if (Number.isNaN(n)) return total;

      // Add the value (use Math.trunc if you only want integer part)
      return total + n;
    }, 0);
  }


  notesView() {
    this.notesViewState = !this.notesViewState
  }
  notesData: any = {}
  Notespopup: any;
  scrollpositionstoring: any;
  addNotes(data: any, ref: any) {
    // this.scrollpositionstoring = this.scrollCurrentposition
    // this.notesData = {
    //   store: data.dealerid,
    //   mainkey: data.dealid,
    //   module: 'CD',
    //   notes: data.Notes,
    //   apiRoute: 'AddGeneralNotes'
    // }

    this.notesData = {
      store: data.dealerid,
      title1: data.dealid,
      title2: '',
      apiRoute: 'AddGeneralNotes'
    }
    this.Notespopup = this.shared.ngbmodal.open(ref, { size: 'lg', backdrop: 'static' });
  }
  closeNotes(e: any) {
    this.shared.ngbmodal.dismissAll()
    if (e == 'S') {
      this.details = []
      this.cardealsdata = [];
      this.callLoadingState = 'ANS'
      this.GetData()
    }
  }
  isDesc: boolean = false;
  column: string = 'CategoryName';

  tradeDetails: any = []
  spinnerLoaderTrade: boolean = false;
  openTradePopup(item: any) {
    this.tradeDetails = []
    this.spinnerLoaderTrade = true
    const obj = {
      "startdealdate": this.FromDate,
      "enddealdate": this.ToDate,
      "dealid": item.dealid
    }
    this.shared.api.postmethod(this.shared.common.routeEndpoint + 'GetSalesCarDealsTradeDetails', obj).subscribe((res: any) => {
      if (res.status == 200) {
        if (res.response != undefined) {
          if (res.response.length > 0) {
            this.tradeDetails = res.response
            this.spinnerLoaderTrade = false
          } else {
            this.spinnerLoaderTrade = false
          }
        } else {
          this.spinnerLoaderTrade = false
        }
      } else {
        this.spinnerLoaderTrade = false
      }
    })
  }
  sort(property: any) {
    this.isDesc = !this.isDesc; //change the direction
    this.column = property;
    let direction = this.isDesc ? 1 : -1;

    this.cardealsdata.sort((a: any, b: any) => {
      const valA = a[property] ?? ''; // replace null/undefined with empty string
      const valB = b[property] ?? '';

      if (valA < valB) {
        return -1 * direction;
      } else if (valA > valB) {
        return 1 * direction;
      } else {
        return 0;
      }
    });

  }
  subdataindex: any = 0;
  excel!: Subscription;

  async openSalesModal(dealnumber: any, vin: any, storeid: any, stock: any, source: any, custno: any) {
    const module = await import('../../../../Layout/cdpdataview/deal/deal-module');
    const component = module.Deal;

    const modalRef = this.shared.ngbmodal.open(component, { size: 'xl', windowClass: 'connectedmodal' });
    modalRef.componentInstance.data = { dealno: dealnumber, vin: vin, storeid: storeid, stock: stock, source: source, custno: custno }; // Pass data to the modal component    
    modalRef.result.then((result) => {
      // this.topScroll()
      console.log(result); // Handle modal close result
    }, (reason) => {
      // this.topScroll()
      console.log(`Dismissed: ${reason}`); // Handle dismiss reason
    });
  }

  ngAfterViewInit() {

    this.shared.api.GetReports().subscribe((data) => {
      let count = data.obj.count
      if (data.obj.Reference == 'Car Deals') {
        count++
        console.log(data.obj);
        this.cardealsdata = [];
        this.pageNumber = 0
        if (data.obj.header == undefined) {
          this.TotalReport = data.obj.TotalReport;
          this.FromDate = data.obj.FromDate;
          this.ToDate = data.obj.ToDate;
          this.store = data.obj.storeValues;
          this.groups = data.obj.groups

          this.dealType = data.obj.dealType;
          this.saleType = data.obj.saleType;
          this.dealStatus = data.obj.dealStatus;
          this.otherblocks = data.obj.otherblock;
          this.obTemporary = data.obj.otherblock
          this.DateType = data.obj.dateType
          this.pageNumber = 0;
          this.QISearchName = data.obj.search;
          this.callLoadingState = 'FL'
          this.setHeaderReportData();

        }
        // this.GetData();
        else {
          if (data.obj.header == 'Yes' && data.obj.Reference == 'Car Deals') {
            count = 0
            this.store = data.obj.storeValues;
            this.pageNumber = 0
            console.log(data.obj);
            // this.GetData();
            this.setHeaderReportData()
          }
        }
      }

    });

    this.excel = this.shared.api.getExportToExcelAllReports().subscribe((res) => {
      if (this.excel != undefined) {
        if (res.obj.title == 'Car Deals') {
          if (res.obj.state == true) {
            this.exportToExcel();
          }
        }
      }
    });


  }

  setHeaderReportData() {
    this.pageNumber = 0

    this.details = [];
    this.cardealsdata = []
    this.obTemporary = this.otherblocks[0]
    const headerdata = {
      title: 'Car Deals',
      path1: this.path1name,
      path2: this.path2name,
      path3: this.path3name,
      path1id: this.path1id,
      path2id: this.path2id,
      path3id: this.path3id,
      stores: this.store,
      dealType: this.dealType,
      saleType: this.saleType,
      dealStatus: this.dealStatus,
      ToporBottom: this.TotalReport,
      fromdate: this.FromDate,
      todate: this.ToDate,
      groups: this.groups,
      otherblock: this.otherblocks,
      search: this.QISearchName,
      count: this.count, datevaluetype: this.DateType
    };
    this.shared.api.SetHeaderData({
      obj: headerdata,
    });
    this.header = [{
      type: 'Bar', storeIds: this.store, fromDate: this.FromDate, toDate: this.ToDate, ReportTotal: this.TotalReport, search: this.QISearchName, groups: this.groups, datevaluetype: this.DateType, dealType: this.dealType,
      saleType: this.saleType,
      dealStatus: this.dealStatus,
    }]
    if (this.store != '') {
      this.GetData()
    } else {
      this.NoData = true;
      this.cardealsdata = []
    }
  }
  ExcelStoreNames: any = [];

  exportToExcel(): void {
    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('Car Deals');

    const PresentYear = this.shared.datePipe.transform(this.FromDate, 'yyyy');
    const FromDate = this.shared.datePipe.transform(this.FromDate, 'dd');
    const ToDate = this.shared.datePipe.transform(this.ToDate, 'dd');
    const PresentMonth = this.shared.datePipe.transform(this.FromDate, 'MMMM');
    this.ExcelStoreNames = []
    let storeNames: any[] = [];
    let store: any = [];
    this.store && this.store.toString().indexOf(',') > 0 ? store = this.store.toString().split(',') : store.push(this.store)
    console.log(this.store, this.groups, '...............................');

    storeNames = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groups)[0].Stores.filter((item: any) => store.includes(item.ID.toString()));
    if (store.length == this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groups)[0].Stores.length) { this.ExcelStoreNames = 'All Stores' }
    else { this.ExcelStoreNames = storeNames.map(function (a: any) { return a.storename; }); }

    //     let storeValue = 'All Stores';
    // if (
    //   this.storeIds &&
    //   this.storeIds.length > 0 &&
    //   this.storeIds.length !== this.stores.length
    // ) {
    //   storeValue = this.stores
    //     .filter((s: any) => this.storeIds.includes(s.ID))
    //     .map((s: any) => s.storename)
    //     .join(', ');
    // }

    let filters: any = [
      // { name: 'Store :', values: 'WesternAuto' },
      { name: 'Time Frame :', values: this.FromDate + ' to ' + this.ToDate },
      { name: 'New Used : ', values: this.dealType == '' ? '-' : this.dealType == null ? '-' : this.dealType.toString().replaceAll(',', ', ') },
      { name: 'Deal Type :', values: this.saleType == '' ? '-' : this.saleType == null ? '-' : this.saleType.toString().replaceAll(',', ', ').replace('Rental', 'Rental/Loaner') },
      { name: 'Deal Status :', values: this.dealStatus == '' ? '-' : this.dealStatus == null ? '-' : this.dealStatus.toString().replaceAll(',', ', ').replace('Capped', 'Booked').replace('Finalized', 'Finalized') },
    ]
    // const ReportFilter = worksheet.addRow(['Report Controls :']);
    // ReportFilter.font = { name: 'Arial', family: 4, size: 10, bold: true };

    const titleRow = worksheet.addRow(['Car Deals']);
    titleRow.eachCell((cell, number) => {
      cell.alignment = {
        indent: 1,
        vertical: 'middle',
        horizontal: 'left',
      };
    });
    titleRow.font = { name: 'Arial', family: 4, size: 12, bold: true };
    titleRow.worksheet.mergeCells('A2', 'D2');

    const Stores1 = worksheet.getCell('A3');
    Stores1.value = 'Stores :';
    worksheet.mergeCells('B3', 'Z3');
    const stores1 = worksheet.getCell('B3');
    stores1.value = this.ExcelStoreNames.toString().replaceAll(',', ', ');
    stores1.font = { name: 'Arial', family: 4, size: 9 };
    stores1.alignment = { vertical: 'top', horizontal: 'left', wrapText: true, };

    let startIndex = 3
    filters.forEach((val: any) => {
      startIndex++
      worksheet.addRow('');
      worksheet.getCell(`A${startIndex}`);
      worksheet.mergeCells(`B${startIndex}:C${startIndex}`);
      worksheet.getCell(`A${startIndex}`).value = val.name;
      worksheet.getCell(`B${startIndex}`).value = val.values
    })
    var secondHeader = []
    var bindingHeaders = []
    if (storeNames && storeNames.length > 1) {
      secondHeader = [
        'Date', 'Store', 'Deal #', 'Status', 'R/L', 'Stock #', 'New/Used', 'Year', 'Make', 'Model', 'Vehicle Age', 'Trade',
        'VIN', 'Age', 'F Gross', 'B Gross', 'Total', 'Mgr', 'F&I Mgr', 'Sp 1', 'Sp 2', 'Buyer', 'Cust #'
      ];
      bindingHeaders = [
        'displaydate', 'store', 'dealid', 'Status', 'RL', 'Stock', 'Type',
        'ad_year', 'ad_make', 'ad_model', 'VehicleAge', 'TradeACV', 'VIN',
        'ad_custAge', 'frontgross', 'backgross', 'totalgross', 'salesmanager',
        'fimanager', 'salesperson1', 'salesperson2', 'Buyer', 'ad_custid',
      ];
    } else {
      secondHeader = [
        'Date', 'Deal #', 'Status', 'R/L', 'Stock #', 'New/Used', 'Year', 'Make', 'Model', 'Vehicle Age', 'Trade',
        'VIN', 'Age', 'F Gross', 'B Gross', 'Total', 'Mgr', 'F&I Mgr', 'Sp 1', 'Sp 2', 'Buyer', 'Cust #'
      ];
      bindingHeaders = [
        'displaydate', 'dealid', 'Status', 'RL', 'Stock', 'Type',
        'ad_year', 'ad_make', 'ad_model', 'VehicleAge', 'TradeACV', 'VIN',
        'ad_custAge', 'frontgross', 'backgross', 'totalgross', 'salesmanager',
        'fimanager', 'salesperson1', 'salesperson2', 'Buyer', 'ad_custid',
      ];
    }

    worksheet.addRow(secondHeader);
    worksheet.getRow(8).height = 25;
    worksheet.getRow(8).font = { bold: true, color: { argb: 'FFFFFFFF' } };
    worksheet.getRow(8).alignment = { vertical: 'middle', horizontal: 'center' };
    worksheet.getRow(8).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2F5597' } };


    const capitalize = (str: string) =>
      str ? str.toString().replace(/\b\w/g, char => char.toUpperCase()) : '';



    const currencyFields = ['TradeACV', 'frontgross', 'backgross', 'totalgross'];
    for (const info of this.cardealsdata) {
      const rowData = bindingHeaders.map(key => {
        const val = info[key];
        // displaydate
        if (key === 'displaydate') return this.shared.datePipe.transform(val, 'MM/dd/yyyy');

        return val === 0 || val == null ? '-' : capitalize(val);
      });

      const dealerRow = worksheet.addRow(rowData);
      // dealerRow.font = { bold: false };

      bindingHeaders.forEach((key, index) => {
        const cell = dealerRow.getCell(index + 1);
        if (currencyFields.includes(key) && typeof cell.value === 'number') {
          cell.numFmt = '"$"#,##0.00';
          cell.alignment = { horizontal: 'right' };
        } else if (!isNaN(Number(cell.value))) {
          cell.alignment = { horizontal: 'center' };
        }
      });
    }
    worksheet.columns.forEach((column: any) => {
      let maxLength = 15;
      column.width = maxLength + 2;
    });
    workbook.xlsx.writeBuffer().then((buffer: any) => {
      const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      this.shared.exportToExcel(workbook, 'Car Deals')
    });
  }



}