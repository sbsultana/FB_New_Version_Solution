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
@Component({
  selector: 'app-dashboard',
  imports: [SharedModule, BsDatepickerModule, DateRangePicker, Stores],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
})
export class Dashboard {
  ServiceData: any = [];
  IndividualServiceGross: any = [];
  TotalServiceGross: any = [];

  TotalReport: string = 'T';
  NoData: any = '';

  responcestatus: any = '';
  currentDate = new Date();
  groups: any = [1];

  reportOpenSub!: Subscription;
  reportGetting!: Subscription;
  Pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;

  type: any = 'C';
  DupType: any = 'C'
  department: any = ['Service', 'Parts', 'Quicklube'];
  includeValues: any = []
  Mileage: any = false;

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
    let today = new Date();
    this.currentDate = new Date(today.setDate(today.getDate() - 1));
    this.shared.setTitle(this.comm.titleName + '-RO Traffic');
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
    this.initializeDates('MTD')

    this.setHeaderData();
    this.getServiceData()
  }

  setHeaderData() {
    const data = {
      title: 'RO Traffic',
      stores: this.storeIds,
      ToporBottom: this.TotalReport,
      fromdate: this.FromDate,
      todate: this.ToDate,
      groups: this.groups,
      department: this.department

    };
    this.shared.api.SetHeaderData({ obj: data });
    // this.getServiceData()


  }
  getServiceData() {
    if (this.storeIds != '') {
      this.responcestatus = '';
      this.shared.spinner.show();
      this.DupType = this.type
      this.GetData();
      // this.GetTotalData();
    } else {
      this.NoData = true
    }
  }
  initializeDates(type: any) {
    let dates: any = this.setdates.setDates(type)
    this.FromDate = dates[0];
    this.ToDate = dates[1];
    localStorage.setItem('time', type);
    this.DateType = type
    this.setDates(this.DateType)

  }

  includes(e: any) {
    let truevalues: any = [this.Mileage]
    this.includeValues = truevalues.filter((val: any) => val == true)

  }
  GetData() {

    const obj = {
      "StartDate": this.shared.datePipe.transform(this.currentDate,'MM-dd-yyyy'),
      "Count": "2",
      "StoreID": this.storeIds.toString(),
      "PayType": "C,W,I",
      DepartmentS: this.department.indexOf('Service') >= 0 ? 'S' : '',
      DepartmentP: this.department.indexOf('Parts') >= 0 ? 'P' : '',
      // DepartmentP: '',
      DepartmentQ: this.department.indexOf('Quicklube') >= 0 ? 'Q' : '',
      DepartmentB: this.department.indexOf('Body') >= 0 ? 'B' : '',
      DepartmentD: this.department.indexOf('Details') >= 0 ? 'D' : '',
      "flag": "",
      "type": this.type,
      exclude: this.Mileage == true ? 'Y' : '',

      firstdate: this.FromDate.replaceAll('/', '-'),
      lastdate: this.ToDate.replaceAll('/', '-'),
    }
    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetROTrafficDate', obj).subscribe((res: any) => {
      if (res.status == 200) {
        if (res && res.response && res.response.length > 0) {
          this.shared.spinner.hide();
          let individualdata = res.response.filter(
            (e: any) => e.Store.toLowerCase() != 'report totals'
          );
          let Totalsdata = res.response.filter(
            (e: any) => e.Store.toLowerCase() == 'report totals'
          );
          if (this.TotalReport == 'T') {
            this.ServiceData = [...Totalsdata, ...individualdata]
          } else {
            this.ServiceData = [...individualdata, ...Totalsdata]

          }
          // this.ServiceData = res.response;

          console.log(this.ServiceData, 'Service Data');

        } else {
          this.shared.spinner.hide();
          this.NoData = 'No Data Found!!'
        }
      }
    })
  }
  GetTotalData() {

  }



  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    }
    else if (value < 0) {
      return false;
    }
    return true
  }
  ngAfterViewInit() {
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.comm.pageName == 'RO Traffic') {
        if (res.obj.storesData != undefined) {
          this.groupsArray = res.obj.storesData;
          this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
          this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_name : this.groupName = ''
          this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
          this.getStoresandGroupsValues()
        }
      }
    })

    this.excel = this.shared.api.getExportToExcelAllReports().subscribe((res) => {
      if (this.excel != undefined) {
        if (res.obj.title == 'RO Traffic') {
          if (res.obj.state == true) {
            this.exportToExcel();
          }
        }
      }
    });
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res) => {
      if (this.email != undefined) {
        if (res.obj.title == 'RO Traffic') {
          if (res.obj.stateEmailPdf == true) {
            //   this.sendEmailData(res.obj.Email, res.obj.notes, res.obj.from);
          }
        }
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res) => {
      if (this.print != undefined) {
        if (res.obj.title == 'RO Traffic') {
          if (res.obj.statePrint == true) {
            // this.GetPrintData();
          }
        }
      }
    });

    this.Pdf = this.shared.api.getExportToPDFAllReports().subscribe((res) => {
      if (this.Pdf != undefined) {
        if (res.obj.title == 'RO Traffic') {
          if (res.obj.statePDF == true) {
            //  this.generatePDF();
          }
        }
      }
    });
  }
  public getColorClass(value: number | null | undefined): string {
    if (value === null || value === undefined || value === 0) {
      return ''; // No class applied
    }
    return value > 0 ? 'positivebg' : 'negativebg';
  }

  isDesc: boolean = false;
  column: string = 'CategoryName';

  sort(property: any, data: any) {
    this.isDesc = !this.isDesc; //change the direction
    this.column = property;
    let direction = this.isDesc ? 1 : -1;
    // //console.log(property)
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



  Scrollpercent: any = 0;
  scrollCurrentposition: any = 0
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

  // Report popup Code
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
  multipleorsingle(block: any, e: any) {

    if (block == 'Dept') {
      const index = this.department.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.department.splice(index, 1);
        if (this.department.length == 0) {
          this.toast.show('Please Select Atleast One Department', 'warning', 'Warning');
        }
      } else {
        this.department.push(e);
      }
    }

    if (block == 'TB') {
      this.TotalReport = e

    }

    if (block == 'T') {
      this.type = e
    }


  }

  activePopover: number = -1;

  togglePopover(popoverIndex: number) {
    this.activePopover = this.activePopover === popoverIndex ? -1 : popoverIndex;
  }

  viewreport() {
    this.activePopover = -1
    if (this.storeIds.length == 0) {
      this.toast.show('Please Select Atleast One Store', 'warning', 'Warning');
    }
    else if (this.department.length == 0) {
      this.toast.show('Please Select Atleast One Department', 'warning', 'Warning');
    }
    else {
      this.responcestatus = ''
      this.setHeaderData();
      this.getServiceData()
    }
  }
  ExcelStoreNames: any = []
  exportToExcel(): void {
    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('RO Traffic');

    const titleRow = worksheet.addRow(['RO Traffic']);
    titleRow.eachCell((cell, number) => {
      cell.alignment = {
        indent: 1,
        vertical: 'middle',
        horizontal: 'left',
      };
    });
    titleRow.font = { name: 'Arial', family: 4, size: 12, bold: true };
    titleRow.worksheet.mergeCells('A2', 'D2');

    const PresentYear = this.shared.datePipe.transform(this.FromDate, 'yyyy');
    const FromDate = this.shared.datePipe.transform(this.FromDate, 'dd');
    const ToDate = this.shared.datePipe.transform(this.ToDate, 'dd');
    const PresentMonth = this.shared.datePipe.transform(this.FromDate, 'MMMM');

    let store = this.storeIds;
    // let storeNames = this.comm.groupsandstores.filter((v: any) => v.sg_id == this.groups)[0].Stores.filter((item: any) => store.some((cat: any) => cat === item.ID.toString()));

    let group = this.comm.groupsandstores.find((v: any) => v.sg_id === this.groups[0]);
    console.log(group, store, 'Group');

    let storeNames = group ? group.Stores.filter((item: any) => store.includes(item.ID)) : [];

    if (store.length == this.comm.groupsandstores.filter((v: any) => v.sg_id == this.groups)[0].Stores.length) {
      this.ExcelStoreNames = 'All Stores';
    } else {
      this.ExcelStoreNames = storeNames.map(function (a: any) {
        return a.storename.toString();
      });
    }

    const Stores1 = worksheet.getCell('A3');
    Stores1.value = 'Stores :';
    worksheet.mergeCells('B3', 'O5');
    const stores1 = worksheet.getCell('B3');
    stores1.value = this.ExcelStoreNames == 0 ? 'All Stores' : this.ExcelStoreNames == null ? '-' : this.ExcelStoreNames.toString().replaceAll(',', ', ');
    stores1.font = { name: 'Arial', family: 4, size: 9 };
    stores1.alignment = {
      vertical: 'top',
      horizontal: 'left',
      wrapText: true,
    };

    let filters: any = [
      // { name: 'Store :', values: this.ExcelStoreNames },
      { name: 'Depatrment :', values: this.department.toString() },
      { name: 'Type : ', values: this.type == 'C' ? 'Car Count' : 'Gross' },
      { name: 'Exclude :', values: this.Mileage == true ? 'Mileage < 10,000' : '-' },
      { name: 'Report Total :', values: this.TotalReport == 'T' ? 'Top' : 'Bottom' },
    ]
    // const ReportFilter = worksheet.addRow(['Report Controls :']);
    // ReportFilter.font = { name: 'Arial', family: 4, size: 10, bold: true };



    let startIndex = 6
    filters.forEach((val: any) => {
      startIndex++
      worksheet.addRow('');
      worksheet.getCell(`A${startIndex}`);
      worksheet.mergeCells(`B${startIndex}:C${startIndex}`);
      worksheet.getCell(`A${startIndex}`).value = val.name;
      worksheet.getCell(`B${startIndex}`).value = val.values
    })

    worksheet.addRow('');
    worksheet.getCell('A11');
    worksheet.mergeCells('B11:E11');
    worksheet.mergeCells('F11:I11');
    worksheet.mergeCells('J11:M11');
    worksheet.mergeCells('N11:Q11');
    worksheet.mergeCells('R11:U11');
    worksheet.mergeCells('V11:Y11');


    worksheet.getCell('A11').value = ``;
    worksheet.getCell('B11').value = 'Car Count (Except IP)';
    worksheet.getCell('F11').value = 'Warranty Only';
    worksheet.getCell('J11').value = 'Customer Pay';
    worksheet.getCell('N11').value = 'Warranty Pay';
    worksheet.getCell('R11').value = 'Extended Warranty (CSC,CSCT,WESP)';
    worksheet.getCell('V11').value = 'Internal Pay';


    worksheet.getRow(1).height = 25;


    ['A11', 'B11', 'F11', 'J11', 'N11', 'R11', 'V11'].forEach(key => {
      const cell = worksheet.getCell(key);
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2F5597' } };
    });

    const secondHeader = [
      `Store`,
      'Curr Year', 'Prev Year', '+/-', '% Change',
      'Curr Year', '%', 'Prev Year', '%',
      'Curr Year', 'Prev Year', '+/-', '% Change',
      'Curr Year', 'Prev Year', '+/-', '% Change',
      'Curr Year', 'Prev Year', '+/-', '% Change',
      'Curr Year', 'Prev Year', '+/-', '% Change',
    ];

    worksheet.addRow(secondHeader);
    worksheet.getRow(2).font = { bold: true };
    worksheet.getRow(2).alignment = { horizontal: 'center', vertical: 'middle' };

    const bindingHeaders = [
      'Store',
      'CW_CarCount_CY', 'CW_CarCount_PY', 'CW_carcount_diff', 'CW_carcount_percent',
      'W_ROCount_CY', 'current_percent', 'W_ROCount_PY', 'Prev_precent',
      'CP_CURR_YR', 'CP_PREV_YR', 'CP_Diff', 'CP_Percent',
      'W_CURR_YR', 'W_PREV_YR', 'W_diff', 'W_Percent',
      'E_CURR_YR', 'E_PREV_YR', 'E_diff', 'E_Percent',
      'I_CURR_YR', 'I_PREV_YR', 'I_diff', 'I_Percent',

    ];
    var currencyFields: any = [];

    this.type == 'C' ? currencyFields = [
      'CP_CURR_YR', 'CP_PREV_YR', 'CP_Diff',
      'W_CURR_YR', 'W_PREV_YR', 'W_diff',
      'E_CURR_YR', 'E_PREV_YR', 'E_diff',
      'I_CURR_YR', 'I_PREV_YR', 'I_diff'] : currencyFields = []

    var Percent: any = ['CP_Percent', 'W_Percent', 'E_Percent', 'I_Percent', 'CW_carcount_percent', 'current_percent', 'Prev_precent']

    const capitalize = (str: string) =>
      str ? str.toString().replace(/\b\w/g, char => char.toUpperCase()) : '';

    for (const info of this.ServiceData) {
      const rowData = bindingHeaders.map(key => {
        const val = info[key];
        return val === 0 || val == null ? '-' : (Percent.includes(key) ? '      ' + val + '%' : capitalize(val));
      });

      const dealerRow = worksheet.addRow(rowData);
      dealerRow.font = { bold: true };

      bindingHeaders.forEach((key, index) => {
        const cell = dealerRow.getCell(index + 1);
        if (currencyFields.includes(key) && typeof cell.value === 'number') {
          cell.numFmt = '"$"#,##0.00';
          cell.alignment = { horizontal: 'right' };
        }
        else if (!isNaN(Number(cell.value))) {
          cell.alignment = { horizontal: 'right' };
        }
      });

    }
    worksheet.columns.forEach((column: any) => {
      let maxLength = 15;
      column.width = maxLength + 2;
    });

      this.shared.exportToExcel(workbook, 'RO Traffic_' + '.xlsx');

  }



}

