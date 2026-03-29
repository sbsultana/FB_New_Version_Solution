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
  NoData: any = '';
  DupFromDate: any = '';
  DupToDate: any = ''
  Pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;
  LendersData: any = []
  LendersDataDetails: any = []
  LendersIndividual: any = [];
  LendersTotal: any = []
  Lenders: any = []

  dataGrouping: any = [
    { "ARG_ID": 1, "ARG_LABEL": "Store", "ARG_SEQ": 0, 'Column_Name': 'store_name' },
    { "ARG_ID": 2, "ARG_LABEL": "Carrier", "ARG_SEQ": 2, 'Column_Name': 'InsName' },
    // { "ARG_ID": 28, "ARG_LABEL": "Tech Name", "ARG_SEQ": 3, 'Column_Name': 'TechName' },
    // { "ARG_ID": 26, "ARG_LABEL": "Counter Person", "ARG_SEQ": 1, 'Column_Name': 'AP_CounterPerson' },


  ];

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
  constructor(public shared: Sharedservice, public setdates: Setdates, private comm: common, private cp: CurrencyPipe, private toast: ToastService,
  ) {
    this.shared.setTitle(this.comm.titleName + '- Insurance Carrier Analysis');
    if (localStorage.getItem('flag') == 'V') {
      this.storeIds = [];
      console.log(JSON.parse(localStorage.getItem('userInfo')!), JSON.parse(localStorage.getItem('userInfo')!).user_Info, 'Widget Stores............');
      this.groupId = JSON.parse(localStorage.getItem('userInfo')!).groupid
      JSON.parse(localStorage.getItem('userInfo')!).store.indexOf(',') > 0 ?
        this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).store.split(',') :
        this.storeIds.push(JSON.parse(localStorage.getItem('userInfo')!).store)
      localStorage.setItem('flag', 'M')
    } else {
      if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
        this.groupId = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Preferences
        this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids.split(',')
      }
    }
    if (this.shared.common.groupsandstores.length > 0) {
      this.groupsArray = this.shared.common.groupsandstores.filter((val: any) => val.sg_id != this.shared.common.reconID);
      this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
      this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_Name : this.groupName = ''
      this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
      this.getStoresandGroupsValues()
    }
    this.initializeDates('MTD')
    this.grouping.push(this.dataGrouping[0]);
    this.grouping.push(this.dataGrouping[1]);

    this.SetHeaderData();
    this.GetLenders()
  }
  ngOnInit(): void { }

  initializeDates(type: any) {
    let dates: any = this.setdates.setDates(type)
    this.FromDate = dates[0];
    this.ToDate = dates[1];
    localStorage.setItem('time', type);
    this.DateType = type
    this.setDates(this.DateType)

  }
  isDesc: boolean = false;
  column: string = '';
  sort(property: any) {
    this.isDesc = !this.isDesc; //change the direction
    this.column = property;
    let direction = this.isDesc ? 1 : -1;
    // // console.log(property)
    this.LendersData.sort(function (a: any, b: any) {
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
    this.DupToDate = this.ToDate;
    // this.DupGrouping = this.grouping;
    this.DupGrouping = JSON.parse(JSON.stringify(this.grouping));
    this.LendersData = [];
    this.NoData = '';
    this.shared.spinner.show();
    const obj = {
      // "fromDate": this.FromDate,
      // "toDate": this.ToDate,
      // 'Stores': this.storeIds.toString()

      "StartDate": this.FromDate,
      "EndDate": this.ToDate,
      "Stores": this.storeIds.toString(),
      "Dealtype": this.neworused.toString(),
      "Saletype": this.retailorlease.toString(),
      "Dealstatus": this.dealstatus.toString(),
      "Var1": this.grouping.length > 1 ? '' : this.grouping[0].Column_Name,
      // "Var2": this.grouping.length > 1 ? this.grouping[1].Column_Name : '',
    };
    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetInsuranceReportSummary', obj).subscribe(
      (res) => {
        if (res.status == 200) {
          this.LendersData = [];
          if (res.response != undefined) {
            if (res.response.length > 0) {
              // this.LendersTotal = res.response.filter(
              //   (e: any) => e.Location == 'REPORT TOTAL'
              // );
              let data = res.response.filter((val: any) => val?.data2 != 'StoreTotal')
              this.LendersIndividual = data.reduce(
                (r: any, { data1, }: any) => {
                  if (!r.some((o: any) => o.data1 == data1)) {
                    r.push({
                      data1,
                      sub: data.filter((v: any) => v.data1 == data1),

                    });
                  }
                  return r;
                },
                []
              );


              this.LendersIndividual.forEach((val: any) => {

                val.Dealcount = val.sub.reduce((acc: any, curr: any) => {
                  return acc + (curr.Dealcount || 0)
                }, 0)
                val.PercentOfall = val.sub.reduce((acc: any, curr: any) => {
                  return acc + (curr.PercentOfall || 0)
                }, 0)
              })
              this.NoData = '';

              this.LendersData = this.LendersIndividual;
              console.log(this.LendersData);
              this.shared.spinner.hide();
            } else {
              this.shared.spinner.hide();
              this.NoData = 'No Data Found';
            }
          } else {
            this.shared.spinner.hide();
            this.NoData = 'No Data Found';
          }
        } else {
          this.toast.show(res.status, 'danger', 'Error');
          this.shared.spinner.hide();
          this.NoData = 'No Data Found';
        }
      },
      (error) => {
        this.toast.show('502 Bad Gate Way Error', 'danger', 'Error');
        this.shared.spinner.hide();
        this.NoData = 'No data Found';
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

  grouping: any = []
  DupGrouping: any = []
  neworused: any = ['New', 'Used'];
  retailorlease: any = ['Retail', 'Lease', 'Misc'];
  dealTypeData: any = ['All', 'Retail', 'Lease', 'Wholesale', 'Misc', 'Fleet', 'Demo', 'Special Order', 'Rental', 'Dealer Trade']
  dealstatus: any = ['Delivered', 'Capped', 'Finalized'];
  multipleorsingle(block: any, e: any) {

    if (block == 'GRP') {
      const index = this.grouping.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.grouping.splice(index, 1);
      } else {
        if (this.grouping.length == 2) {
          this.toast.show('Select up to 2 Filters only to Group your data', 'warning', 'Warning');
        } else {
          this.grouping.push(e);
        }
      }
    }
    if (block == 'NU') {
      const index = this.neworused.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.neworused.splice(index, 1);
      } else {
        this.neworused.push(e);
      }
      if (this.neworused.length == 0) {
        this.alert('New/Used')
      }
    }
    if (block == 'RL') {
      const index = this.retailorlease.findIndex((i: any) => i == e);
      if (e == 'All') {
        let allindex = this.retailorlease.findIndex((i: any) => i == 'All');
        if (allindex >= 0) {
          this.retailorlease = []
        }
        else {
          const dealdata = JSON.stringify(this.dealTypeData)
          this.retailorlease = JSON.parse(dealdata)
        }
      } else {
        if (index >= 0) {
          this.retailorlease.splice(index, 1);
        } else {
          this.retailorlease.push(e)
        }
      }
      if (this.retailorlease.length == 0) {
        this.alert('Dealtype')
      }
    }
    if (block == 'DS') {
      const index = this.dealstatus.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.dealstatus.splice(index, 1);
      } else {
        this.dealstatus.push(e);
      }
      if (this.dealstatus.length == 0) {
        this.alert('Dealstatus')
      }
    }
  }
  alert(block: any) {
    this.toast.show('Please select atleast any one ' + block, 'warning', 'Warning')
  }

  LMstate: any;
  ngAfterViewInit() {
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.comm.pageName == 'Insurance Carrier Analysis') {
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
        if (res.obj.title == 'Insurance Carrier Analysis') {
          if (res.obj.stateEmailPdf == true) {
            //    this.sendEmailData(res.obj.Email, res.obj.notes, res.obj.from);
          }
        }
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res) => {
      if (this.print != undefined) {
        if (res.obj.title == 'Insurance Carrier Analysis') {
          if (res.obj.statePrint == true) {
            // this.GetPrintData();
          }
        }
      }
    });
    this.Pdf = this.shared.api.getExportToPDFAllReports().subscribe((res) => {
      if (this.Pdf != undefined) {
        if (res.obj.title == 'Insurance Carrier Analysis') {
          if (res.obj.statePDF == true) {
            // this.generatePDF();
          }
        }
      }
    });

    this.excel = this.shared.api.getExportToExcelAllReports().subscribe((res) => {
      if (this.excel != undefined) {
        this.LMstate = res.obj.state;
        if (res.obj.title == 'Insurance Carrier Analysis') {
          if (res.obj.state == true) {
            this.exportToExcel();
          }
        }
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
  SetHeaderData() {
    const data = {
      title: 'Insurance Carrier Analysis',
      path1: this.grouping[0].ARG_LABEL,
      path2: this.grouping.length > 1 ? this.grouping[1].ARG_LABEL : '',
      stores: this.storeIds,
      groups: this.groupId,
      fromdate: this.FromDate,
      todate: this.ToDate,
      dealtype: this.neworused,
      saletype: this.retailorlease,
      dealstatus: this.dealstatus
    };
    this.shared.api.SetHeaderData({
      obj: data,
    });
  }
  GetPrintData() {
    window.print();
  }
  ngOnDestroy() {

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


  activePopover: number = -1;
  togglePopover(popoverIndex: number) {
    if (this.activePopover === popoverIndex) {
      this.activePopover = -1;
    } else {
      this.activePopover = popoverIndex;
    }
  }



  viewreport() {
    this.activePopover = -1
    if (this.storeIds.length == 0) {
      this.toast.show('Please Select Atleast One Store', 'warning', 'Warning');
    }
    else if (this.neworused.length == 0) {
      this.alert('New/Used')
    }
    else if (this.retailorlease.length == 0) {
      this.alert('Dealtype')
    }
    else if (this.dealstatus.length == 0) {
      this.alert('Dealstatus')
    }
    else {
      // const data = {
      //   Reference: 'Insurance Carrier Analysis',
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


  ExcelStoreNames: any = [];
  exportToExcel() {
    let storeNames: any = [];
    let store = this.storeIds
    storeNames = this.comm.groupsandstores
      .find((v: any) => v.sg_id == this.groupId)
      ?.Stores.filter((item: any) => store.includes(Number(item.ID)));
    console.log(storeNames);

    if (store.length == this.comm.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores.length) {
      this.ExcelStoreNames = 'All Stores'
    } else {
      this.ExcelStoreNames = storeNames.map(function (a: any) {
        return a.storename;
      });
      console.log(this.ExcelStoreNames);

    }
    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('Insurance Carrier Analysis');
    worksheet.views = [
      {
        state: 'frozen',
        ySplit: 12, // Number of rows to freeze (2 means the first two rows are frozen)
        topLeftCell: 'A13', // Specify the cell to start freezing from (in this case, the third row)
        showGridLines: false,
      },
    ];
    worksheet.addRow('');
    const titleRow = worksheet.addRow(['Insurance Carrier Analysis']);
    titleRow.eachCell((cell, number) => {
      cell.alignment = { indent: 1, vertical: 'top', horizontal: 'left' };
    });
    titleRow.font = { name: 'Arial', family: 4, size: 12, bold: true };
    titleRow.worksheet.mergeCells('A2', 'D2');
    const PresentMonth = this.shared.datePipe.transform(this.FromDate, 'MMMM');
    const PresentYear = this.shared.datePipe.transform(this.FromDate, 'yyyy');
    const FromDate = this.shared.datePipe.transform(this.FromDate, 'dd');
    const ToDate = this.shared.datePipe.transform(this.ToDate, 'dd');
    worksheet.addRow('');
    const DateToday = this.shared.datePipe.transform(
      new Date(),
      'MM/dd/yyyy h:mm:ss a'
    );
    const DATE_EXTENSION = this.shared.datePipe.transform(new Date(), 'MMddyyyy');
    worksheet.addRow([DateToday]).font = { name: 'Arial', family: 4, size: 9 };
    const Appointmentdata = this.LendersData.map((_arrayElement: any) =>
      Object.assign({}, _arrayElement)
    );
    const ReportFilter = worksheet.addRow(['Report Filters :']);
    ReportFilter.font = { name: 'Arial', family: 4, size: 10, bold: true };
    const Groups = worksheet.getCell('A6');
    Groups.value = 'Group :';
    Groups.font = { name: 'Arial', family: 4, size: 9, bold: true };
    const groups = worksheet.getCell('B6');
    groups.value =
      this.comm.groupsandstores.filter((val: any) => val.sg_id == this.groupId.toString())[0].sg_name;
    groups.font = { name: 'Arial', family: 4, size: 9 };
    groups.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
    worksheet.mergeCells('B7', 'K9');
    const Stores = worksheet.getCell('A7');
    Stores.value = 'Stores :'
    const stores = worksheet.getCell('B7');
    stores.value = this.ExcelStoreNames == 0
      ? '-'
      : this.ExcelStoreNames == null
        ? '-'
        : this.ExcelStoreNames.toString().replaceAll(',', ', ');
    stores.font = { name: 'Arial', family: 4, size: 9 };
    stores.alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    Stores.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
    };
    const Timeframe = worksheet.addRow(['Timeframe :']);
    Timeframe.getCell(1).font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
    };
    const timeframe = worksheet.getCell('B10');
    timeframe.value = this.FromDate + ' to ' + this.ToDate;
    timeframe.font = { name: 'Arial', family: 4, size: 9 };
    worksheet.addRow('');
    // const Note = worksheet.addRow(['Note :']);const note = worksheet.getCell("C4");note.value = this.StoreValues;note.font = {name: 'Arial',family: 4,size: 8,bold: false,};
    // Note.font = {name: 'Arial',family: 4,size: 8,bold: true,};
    let Headings = [
      'Insurance Name ',
      'Deal Count',
      'Percent of All',

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
    const SalesData = this.LendersData.map((_arrayElement: any) =>
      Object.assign({}, _arrayElement)
    );
    for (const d of SalesData) {
      var data1obj = [
        d.data1 == '' ? '-' : d.data1 == null ? '-' : d.data1,
        d.Dealcount == '' ? '-' : d.Dealcount == null ? '-' : parseFloat(d.Dealcount),
        d.PercentOfall == '' ? '-' : d.PercentOfall == null ? '-' : parseFloat(d.PercentOfall) + '%',

      ]
      const Data1 = worksheet.addRow(data1obj);
      Data1.font = { name: 'Arial', family: 4, size: 9 };
      Data1.alignment = { vertical: 'middle', horizontal: 'center' };
      Data1.getCell(1).alignment = {
        indent: 1,
        vertical: 'middle',
        horizontal: 'left',
      };
      Data1.eachCell((cell: any, number: any) => {
        cell.border = { right: { style: 'dotted' } };
        // cell.numFmt = '$#,##0.00';

        // if (number > 1) {
        //   cell.numFmt = '#,##0';
        // }
        if (number == 2) {
          cell.numFmt = '#,##0.00';
        }
        if (number > 1 && data1obj[number] != undefined) {
          if (data1obj[number] < 0) {
            Data1.getCell(number + 1).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } };
          }
        }
      });
      if (d.sub != undefined && this.grouping && this.grouping.length > 1 && d.data1 != 'REPORTS TOTAL') {
        for (const d1 of d.sub) {
          var data2obj = [
            d1.data2 == '' ? '-' : d1.data2 == null ? '-' : '      ' + d1.data2,
            d1.Dealcount == '' ? '-' : d1.Dealcount == null ? '-' : parseFloat(d1.Dealcount),
            d1.PercentOfall == '' ? '-' : d1.PercentOfall == null ? '-' : parseFloat(d1.PercentOfall) + '%',

          ]
          const Data2 = worksheet.addRow(data2obj);
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
            // cell.numFmt = '$#,##0.00';

            if (number > 1) {
              cell.numFmt = '#,##0';
            }
            // if (number > 1 && data2obj[number] != undefined) {
            //   if (data2obj[number] < 0) {
            //     Data1.getCell(number + 1).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } };
            //   }
            // }
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
      }
    }
    worksheet.getColumn(1).width = 35; worksheet.getColumn(1).alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
    worksheet.getColumn(2).width = 15;
    worksheet.getColumn(3).width = 20;
    worksheet.getColumn(4).width = 25;
    worksheet.getColumn(5).width = 15;
    worksheet.getColumn(6).width = 15;
    worksheet.addRow([]);
    workbook.xlsx.writeBuffer().then((data: any) => {
      const blob = new Blob([data], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      this.shared.exportToExcel(
        workbook,
        'Insurance_' + DATE_EXTENSION
      );
    });
  }

}