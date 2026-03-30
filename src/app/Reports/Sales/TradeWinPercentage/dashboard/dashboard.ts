import { Component, OnInit, ViewChild, ElementRef, HostListener } from '@angular/core';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { Setdates } from '../../../../Core/Providers/SetDates/setdates';
import { BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { common } from '../../../../common';
import { DateRangePicker } from '../../../../CommonFilters/date-range-picker/date-range-picker';
import { Subscription } from 'rxjs';
import { environment } from '../../../../../environments/environment';
import { Stores } from '../../../../CommonFilters/stores/stores';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
import { CurrencyPipe } from '@angular/common';
import { NgbModalRef } from '@ng-bootstrap/ng-bootstrap';

@Component({
  selector: 'app-dashboard',
  imports: [SharedModule, BsDatepickerModule, DateRangePicker, Stores],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
})
export class Dashboard {
  NoData: any = false;

  DupFromDate: any = '';
  DupToDate: any = ''
  reportOpenSub!: Subscription;
  reportGetting!: Subscription;
  Pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;
  ServiceAbsorption: any = []
  LendersDataDetails: any = []
  LendersIndividual: any = [];
  LendersTotal: any = []
  Lenders: any = []

  toporbottom: any = ['T']

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

  FromDate: any = '';
  ToDate: any = '';
  minDate!: Date;
  maxDate!: Date;
  DateType: any = 'MTD';
  displaytime: any = '';


  Dates: any = {
    'FromDate': this.FromDate, 'ToDate': this.ToDate, "MaxDate": this.maxDate, 'MinDate': this.minDate, 'DateType': this.DateType, 'DisplayTime': this.displaytime,
    Types: [
      { 'code': 'MTD', 'name': 'MTD' },
      { 'code': 'QTD', 'name': 'QTD' },
      { 'code': 'YTD', 'name': 'YTD' },
      { 'code': 'PYTD', 'name': 'PYTD' },
      { 'code': 'LY', 'name': 'Last Year' },
      { 'code': 'LM', 'name': 'Last Month' },
      { 'code': 'PM', 'name': 'Same Month PY' },
    ]
  }


  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent) {
    const clickedInside = (event.target as HTMLElement).closest('.dropdown-toggle, .reportstores-card, .timeframe');
    if (!clickedInside) {
      this.activePopover = -1;
    }
  }
  constructor(
    public shared: Sharedservice, public setdates: Setdates, private comm: common, private cp: CurrencyPipe, private toast: ToastService,
  ) {


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
    this.shared.setTitle(this.comm.titleName + '- Trade Win Percentage');
    // this.setDates('MTD');
    // this.storeIds = this.commSrvc.StoresUserDetails('All', 0).toString();


    this.initializeDates('MTD')

    this.SetHeaderData();
    this.GetLenders();

  }
  ngOnInit(): void {

  }

  initializeDates(type: any) {
    let dates: any = this.setdates.setDates(type)
    this.FromDate = dates[0];
    this.ToDate = dates[1];
    this.DupFromDate = this.FromDate;
    this.DupToDate = this.ToDate
    localStorage.setItem('time', type);
    this.setDates(this.DateType)

  }
  isDesc: boolean = false;
  column: string = '';
  sort(property: any) {
    this.isDesc = !this.isDesc; //change the direction
    this.column = property;
    let direction = this.isDesc ? 1 : -1;
    // // console.log(property)
    this.ServiceAbsorption.sort(function (a: any, b: any) {
      if (a[property] < b[property]) {
        return -1 * direction;
      } else if (a[property] > b[property]) {
        return 1 * direction;
      } else {
        return 0;
      }
    });
  }
  GetLenders() {
    console.log(this.FromDate, this.ToDate);
    this.DupFromDate = this.FromDate;
    this.DupToDate = this.ToDate
    this.ServiceAbsorption = [];
    this.NoData = false;
    this.shared.spinner.show();
    const obj =
    // {
    //   "StartDate": this.FromDate,
    //   "EndDate": this.ToDate,
    //   "Stores": this.storeIds.toString()
    // };
    {
      "StoreIds": this.storeIds.toString(),
      "StartDate": this.FromDate,
      "EndDate": this.ToDate,
      "Dealtype": this.neworused.toString(),
      "DealStatuses": this.dealstatus.toString(),
      "DealType2s": this.retailorlease.toString()
    }

    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetDealTradeWin', obj).subscribe(
      (res) => {
        if (res.status == 200) {
          this.ServiceAbsorption = [];
          if (res.response != undefined) {
            console.log(res.response, 'Response');

            if (res.response.length > 0) {
              let data = res.response.filter((v: any) => v.storeid != 0);
              let total = res.response.filter((v: any) => v.storeid == 0);


              this.toporbottom.toString() == 'T' ? this.ServiceAbsorption = [...total, ...data] : this.ServiceAbsorption = [...data, ...total]
              console.log(this.ServiceAbsorption);
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
          this.toast.show(res.status, 'danger','Error');
          this.shared.spinner.hide();
          this.NoData = true;
        }
      },
      (error) => {
        this.toast.show('502 Bad Gate Way Error', 'danger','Error');
        this.shared.spinner.hide();
        this.NoData = true;
      }
    );
  }

openSalesModal(){

}

  LMstate: any;
  ngAfterViewInit() {

    this.shared.api.getStores().subscribe((res: any) => {
      if (this.comm.pageName == 'Trade Win Percentage') {
        if (res.obj.storesData != undefined) {
          this.groupsArray = res.obj.storesData;
          this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
          this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_name : this.groupName = ''
          this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
          this.getStoresandGroupsValues()
        }
      }
    })

    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res) => {
      if (this.email != undefined) {
        if (res.obj.title == 'Trade Win Percentage') {
          if (res.obj.stateEmailPdf == true) {
          //  this.sendEmailData(res.obj.Email, res.obj.notes, res.obj.from);
          }
        }
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res) => {
      if (this.print != undefined) {
        if (res.obj.title == 'Trade Win Percentage') {
          if (res.obj.statePrint == true) {
          //  this.GetPrintData();
          }
        }
      }
    });
    this.Pdf = this.shared.api.getExportToPDFAllReports().subscribe((res) => {
      if (this.Pdf != undefined) {
        if (res.obj.title == 'Trade Win Percentage') {
          if (res.obj.statePDF == true) {
          //  this.generatePDF();
          }
        }
      }
    });

    this.excel = this.shared.api.getExportToExcelAllReports().subscribe((res) => {
      if (this.excel != undefined) {
        this.LMstate = res.obj.state;
        if (res.obj.title == 'Trade Win Percentage') {
          if (res.obj.state == true) {
            this.exportToExcel();
          }
        }
      }
    });
  }

  SetHeaderData() {
    const data = {
      title: 'Trade Win Percentage',
    
      stores: this.storeIds,
      groups: this.groupId,
      fromdate: this.FromDate,
      todate: this.ToDate,
    };
    this.shared.api.SetHeaderData({
      obj: data,
    });
  }
  GetPrintData() {
    window.print();
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
  updatedDates(data: any) {
    // console.log(data);
    this.FromDate = data.FromDate;
    this.ToDate = data.ToDate;
    this.DateType = data.DateType;
    this.displaytime = data.DisplayTime
  }

  setDates(type: any) {
    this.displaytime = '(' + this.Dates.Types.filter((val: any) => val.code == type)[0].name + ')';
    this.maxDate = new Date();
    this.minDate = new Date();
    this.minDate.setFullYear(this.maxDate.getFullYear() - 3);
    this.maxDate.setDate(this.maxDate.getDate());
    this.Dates.FromDate = this.FromDate;
    this.Dates.ToDate = this.ToDate;
    this.Dates.MinDate = this.minDate;
    this.Dates.MaxDate = this.maxDate;
    this.Dates.DateType = this.DateType;
    this.Dates.DisplayTime = this.displaytime;
  }

  activePopover: number = -1;

  togglePopover(popoverIndex: number) {
    if (this.activePopover === popoverIndex) {
      this.activePopover = -1;
    } else {
      this.activePopover = popoverIndex;
    }
  }


  neworused: any = ['New', 'Used'];
  retailorlease: any = ['Retail', 'Lease', 'Misc'];
  dealstatus: any = ['Delivered', 'Capped', 'Finalized'];

  multipleorsingle(block: any, e: any) {
    if (block == 'NU') {
      const index = this.neworused.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.neworused.splice(index, 1);
      } else {
        this.neworused.push(e);
      }
    }
    if (block == 'RL') {
      const index = this.retailorlease.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.retailorlease.splice(index, 1);
      } else {
        this.retailorlease.push(e);
      }
    }

    if (block == 'DS') {
      const index = this.dealstatus.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.dealstatus.splice(index, 1);
      } else {
        this.dealstatus.push(e);
      }
    }

    if (block == 'TB') {
      this.toporbottom = [];
      this.toporbottom.push(e);
    }

  }

  viewreport() {
    this.activePopover = -1
    if (this.storeIds.length == 0 || this.neworused.length == 0 || this.retailorlease.length == 0 || this.dealstatus.length == 0) {
      if (this.storeIds.length == 0) {
        this.toast.show('Please Select Atleast One Store','warning','Warning');
      }
      if (this.neworused.length == 0) {
        this.toast.show('Please Select Atleast One From New or Used','warning','Warning');
      }
      if (this.retailorlease.length == 0) {
        this.toast.show('Please Select Atleast One Deal Type','warning','Warning');
      }
      if (this.dealstatus.length == 0) {
        this.toast.show('Please Select Atleast One Deal Status','warning','Warning');
      }
    }
    else {
      // const data = {
      //   Reference: 'Trade Win Percentage',
      //   storeValues: this.storeIds.toString(),
      //   groups: this.groups.toString(),
      // };
      // this.shared.api.SetReports({
      //   obj: data,
      // });
      this.SetHeaderData()
      this.GetLenders()
    }
  }


loader:boolean=false
  SalesPersonDetailsNoData: any = ''
  SalesPersonDetails: any = []
  popupReference!: NgbModalRef;
  popupvalues: any = { Acctno: '', AcctDesc: '', Dept: '', Store: '', Storename: '', subtype: '' }
  openDetails(data: any, temp: any, block: any, trade: any) {
    // this.shared.spinner.show()
    this.popupvalues.Store = data.dealername
    // this.popupvalues.Storename = StoreName;
    this.loader=true
    this.popupvalues.subtype = block == 'New,Used' ? 'Totals' : block
    this.SalesPersonDetailsNoData = '';
    this.SalesPersonDetails = []
    this.popupReference = this.shared.ngbmodal.open(temp, { size: 'xl', backdrop: 'static', keyboard: true, centered: true, modalDialogClass: 'custom-modal' })
    const obj = {
      "store": data.storeid,
      "startdealdate": this.FromDate,
      "enddealdate": this.ToDate,
      "dealtype": block,
      "saletype": this.retailorlease.toString(),
      "dealstatus": this.dealstatus.toString(),
      "AcquisitionSource": "",
      "PageNumber": "0",
      "PageSize": "10000000",
      "Trades": trade
    };
    console.log(obj);

    this.shared.api
      .postmethod(this.comm.routeEndpoint + 'GetSalesGrossSummaryDetailsTradeWin', obj)
      .subscribe((res) => {
    this.loader=false
        
        if (res.status == 200) {
          // this.shared.spinner.hide()
          if (res && res.response && res.response.length > 0) {
            this.SalesPersonDetails = res.response;
            this.SalesPersonDetailsNoData = ''

          } else {
            this.SalesPersonDetails = []
            this.SalesPersonDetailsNoData = 'No Data Found!!'
          }
        }
      });
  }
  sortColumn(property: any) {
    this.isDesc = !this.isDesc; //change the direction
    this.column = property;
    let direction = this.isDesc ? 1 : -1;
    // //console.log(direction);

    this.SalesPersonDetails.sort(function (a: any, b: any) {
      if (a[property] < b[property]) {
        return -1 * direction;
      } else if (a[property] > b[property]) {
        return 1 * direction;
      } else {
        return 0;
      }
    });
  }

  closePopup() {
    if (this.popupReference) {
      this.popupReference.close();
    }
  }
  ExcelStoreNames: any = []
  exportToExcel(): void {
    this.shared.spinner.show()
    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('Trade Win Percentage');

    const PresentYear = this.shared.datePipe.transform(this.FromDate, 'yyyy');
    const FromDate = this.shared.datePipe.transform(this.FromDate, 'MM.dd.yyyy');
    const ToDate = this.shared.datePipe.transform(this.ToDate, 'MM.dd.yyyy');
    const PresentMonth = this.shared.datePipe.transform(this.ToDate, 'MMMM');

    let filters: any = [
      { name: 'Time Frame :', values: FromDate + ' - ' + ToDate },
      { name: 'New Used : ', values: this.neworused == '' ? '-' : this.neworused == null ? '-' : this.neworused.toString().replaceAll(',', ', ') },
      { name: 'Deal Type :', values: this.retailorlease == '' ? '-' : this.retailorlease == null ? '-' : this.retailorlease.toString().replaceAll(',', ', ') },
      { name: 'Deal Status :', values: this.dealstatus == '' ? '-' : this.dealstatus == null ? '-' : this.dealstatus.toString().replaceAll(',', ', ').replace('Capped', 'Booked') },
    ]
    // const ReportFilter = worksheet.addRow(['Report Controls :']);
    // ReportFilter.font = { name: 'Arial', family: 4, size: 10, bold: true };
    const titleRow = worksheet.addRow(['Trade Win Percentage']);
    titleRow.eachCell((cell, number) => {
      cell.alignment = {
        indent: 1,
        vertical: 'middle',
        horizontal: 'left',
      };
    });
    titleRow.font = { name: 'Arial', family: 4, size: 12, bold: true };
    titleRow.worksheet.mergeCells('A2', 'D2');


    this.ExcelStoreNames = []
    let storeNames: any[] = [];
    let store: any = [];
    this.storeIds && this.storeIds.toString().indexOf(',') > 0 ? store = this.storeIds.toString().split(',') : store.push(this.storeIds.toString())

    storeNames = this.comm.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores.filter((item: any) => store.includes(item.ID.toString()));
    if (store.length == this.comm.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores.length) { this.ExcelStoreNames = 'All Stores' }
    else { this.ExcelStoreNames = storeNames.map(function (a: any) { return a.storename; }); }
    console.log(this.ExcelStoreNames, 'Store names');

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

    worksheet.addRow('');

    const secondHeader = [
      `${FromDate} - ${ToDate}, ${PresentYear}`, 'NEW', 'TRADES', 'TW%', 'USED', 'TRADES', 'TW%', 'TOTAL', 'TRADES', 'TW%',

    ];

    const headerRow = worksheet.addRow(secondHeader);

    headerRow.eachCell((cell) => {
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.font = { bold: true };
    });

    const bindingHeaders = [
      "dealername", "New_Deals", "New_trades", "TW_new", "Used_deals", "Used_trades", "TW_Used", "Total", "totaltrades", "TW_Total"
    ];
    const currencyFields: any = [];

    const capitalize = (str: string) =>
      str ? str.toString().replace(/\b\w/g, char => char.toUpperCase()) : '';
    for (const info of this.ServiceAbsorption) {
      const rowData = bindingHeaders.map(key => {
        const val = info[key];
        if (key === 'dealername') return val
        if (key === 'TW_new' || key === 'TW_Used' || key === 'TW_Total') return this.cp.transform(val, 'USD', '', '1.0-0') + '%';
        return val === 0 || val == null ? '-' : val
      });

      const dealerRow = worksheet.addRow(rowData);
      dealerRow.font = { bold: true };

      bindingHeaders.forEach((key, index) => {
        const cell = dealerRow.getCell(index + 1);
        if (currencyFields.includes(key) && typeof cell.value === 'number') {
          cell.numFmt = '"$"#,##0';
          cell.alignment = { horizontal: 'right' };
        } else if (!isNaN(Number(cell.value))) {
          cell.alignment = { horizontal: 'right' };
        }
      });


    }
    worksheet.columns.forEach((column: any) => {
      let maxLength = 20;
      column.width = maxLength + 2;
    });
    this.shared.spinner.hide()
  
      this.shared.exportToExcel(workbook, `Trade Win Percentage_${new Date().getTime()}.xlsx`);
  }


  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    } else if (value < 0) {
      return false;
    }
    return true;
  }
}
