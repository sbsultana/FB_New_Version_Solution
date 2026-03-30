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
  SalesContest: any = [];
  NoData: any = false;
  searchQuery: any = ''



  keys: any = [];
  AsofNow: any = ''
  key: any = 'Rank'
  order: any = 'desc'

  reportOpenSub!: Subscription;
  reportGetting!: Subscription;
  Pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;

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
  DupDateType: any = 'MTD'

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

  filterData: any = [
    { 'name': 'Units', 'id': 1 },
    { 'name': 'Gross', 'id': 2 },
  ]

  filtertype: any = ['Units']

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
    this.initializeDates('MTD')
    this.shared.setTitle(this.comm.titleName + '- Sales Contest');
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

    this.setDates(this.DateType)
    this.GetSalesContest();
    this.setHeaderData()
  }
  count: any = 0;


  ngOnInit(): void {
    // var curl = 'https://fbxtract.axelautomotive.com/cavender/GetVSAppointmentsData';
    // this.shared.api.logSaving(curl, {}, '', 'Success', 'Sales Contest');
  }

  initializeDates(type: any) {
    let dates: any = this.setdates.setDates(type)
    this.FromDate = dates[0];
    this.ToDate = dates[1];
    console.log(this.FromDate, this.ToDate);

    localStorage.setItem('time', type);
  }

  sort(key: any, order: any) {
    if (this.key == key) {
      if (order == 'asc') {
        this.order = 'desc'
      } else {
        this.order = 'asc'
      }
    } else {
      this.order = 'asc'
    }
    this.key = key
    // this.order = order
    this.count = 1;
    this.GetSalesContest()
    // this.GetInventorySummaryReport()
  }

  setHeaderData() {
    const data = {
      title: 'Sales Contest',
      stores: this.storeIds,
      groups: this.groupId,
      filtertype: this.filtertype,
      fromdate: this.FromDate,
      todate: this.ToDate,
    };
    this.shared.api.SetHeaderData({
      obj: data,
    });
  }

  datetype() {
    if (this.DupDateType == 'PM') {
      return 'SP';
    }
    else if (this.DupDateType == 'C') {
      return 'C'
    }
    return this.DupDateType;
  }
  isDesc: boolean = false;
  column: string = '';

  TotalData: any = []
  exceldata: any = []
  GetSalesContest() {
    this.SalesContest = [];
    this.NoData = false;
    this.shared.spinner.show();
    this.dupFiltertype = this.filtertype
    this.DupDateType = this.DateType
    const obj = {
      "Stores": this.storeIds,
      "StartDate": this.FromDate,
      "EndDate": this.ToDate,
      "ContestType": this.filtertype.toString(),
      "Exp": this.key,
      "OrderType": this.order,
    };
    const curl = environment.apiUrl + this.comm.routeEndpoint + 'GetSalesContestV1';
    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetSalesContestV1', obj).subscribe(
      (res) => {
        const currentTitle = document.title;
        this.shared.api.logSaving(curl, {}, '', res.message, currentTitle);
        if (res.status == 200) {
          this.TotalData = []
          this.shared.spinner.hide();
          if (res.response != undefined) {
            if (res.response.length > 0) {
              let data = res.response
              this.exceldata = res.response
              this.SalesContest = data.filter((e: any) => e.StoreName != 'Totals');
              this.TotalData = data.filter((e: any) => e.StoreName == 'Totals')
              console.log(this.TotalData);

              let key = Object.keys(res.response[0]);
              this.keys = key.splice(3)
              this.AsofNow = res.response[0].ASofTime
              this.NoData = false;
            } else {
              this.NoData = true;
            }
          } else {
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


  ngAfterViewInit() {
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.comm.pageName == 'Sales Contest') {
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
        if (res.obj.title == 'Sales Contest') {
          if (res.obj.stateEmailPdf == true) {
            // this.sendEmailData(res.obj.Email, res.obj.notes, res.obj.from);
          }
        }
      }
    });

    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res) => {
      if (this.print != undefined) {
        if (res.obj.title == 'Sales Contest') {
          if (res.obj.statePrint == true) {
            // this.GetPrintData();
          }
        }
      }
    });
    this.Pdf = this.shared.api.getExportToPDFAllReports().subscribe((res) => {
      if (this.Pdf != undefined) {
        if (res.obj.title == 'Sales Contest') {
          if (res.obj.statePDF == true) {
            // this.generatePDF();
          }
        }
      }
    });

    this.excel = this.shared.api.getExportToExcelAllReports().subscribe((res) => {
      if (this.excel != undefined) {
        if (res.obj.title == 'Sales Contest') {
          if (res.obj.state == true) {
            this.exportToExcel();
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

  filterselection(e: any) {
    const index = this.filtertype.findIndex((i: any) => i == e.name);
    if (index >= 0) {
      this.filtertype.splice(index, 1);
    } else {
      this.filtertype = []
      this.filtertype.push(e.name);
    }

  }
  dupFiltertype: any = ['Units']

  viewreport() {
    this.activePopover = -1

    if (this.storeIds.length == 0 || this.filtertype.length == 0) {
      if (this.storeIds.length == 0) {
        this.toast.show('Please Select Atleast One Store', 'warning', 'Warning');
      }
      if (this.filtertype.length == 0) {
        this.toast.show('Please Select Atleast One Type', 'warning', 'Warning');
      }
    } else {
      this.setHeaderData()
      this.GetSalesContest()
    }
    // }

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


  ExcelStoreNames: any = [];

  exportToExcel() {
    let storeNames: any = [];
    let store = this.storeIds.split(',');
    storeNames = this.comm.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores.filter((item: any) =>
      store.some((cat: any) => cat === item.ID.toString())
    );
    if (store.length == this.comm.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores.length) {
      this.ExcelStoreNames = 'All Stores'
    } else {
      this.ExcelStoreNames = storeNames.map(function (a: any) {
        return a.storename;
      });
    }
    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('Sales Contest');
    worksheet.views = [
      {
        state: 'frozen',
        ySplit: 11, // Number of rows to freeze (2 means the first two rows are frozen)
        topLeftCell: 'A12', // Specify the cell to start freezing from (in this case, the third row)
        showGridLines: false,
      },
    ];
    worksheet.addRow('');
    const titleRow = worksheet.addRow(['Sales Contest']);
    titleRow.eachCell((cell, number) => {
      cell.alignment = { indent: 1, vertical: 'top', horizontal: 'left' };
    });
    titleRow.font = { name: 'Arial', family: 4, size: 12, bold: true };
    titleRow.worksheet.mergeCells('A2', 'D2');

    worksheet.addRow('');
    const DateToday = this.shared.datePipe.transform(
      new Date(),
      'MM/dd/yyyy h:mm:ss a'
    );
    const DATE_EXTENSION = this.shared.datePipe.transform(new Date(), 'MMddyyyy');
    worksheet.addRow([DateToday]).font = { name: 'Arial', family: 4, size: 9 };
    const Appointmentdata = this.SalesContest.map((_arrayElement: any) =>
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
    const timeframe = worksheet.getCell('B8');
    timeframe.value = this.FromDate + ' to ' + this.ToDate;
    timeframe.font = { name: 'Arial', family: 4, size: 9 };


    const ContestType = worksheet.addRow(['Contest Type:']);
    ContestType.getCell(1).font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
    };
    const contesttype = worksheet.getCell('B9');
    contesttype.value = this.FromDate + ' to ' + this.ToDate;
    contesttype.font = { name: 'Arial', family: 4, size: 9 };
    worksheet.addRow('');
    // const Note = worksheet.addRow(['Note :']);const note = worksheet.getCell("C4");note.value = this.StoreValues;note.font = {name: 'Arial',family: 4,size: 8,bold: false,};
    // Note.font = {name: 'Arial',family: 4,size: 8,bold: true,};

    let Headings = ['Rank', 'Store', this.datetype() == 'C' ? this.shared.datePipe.transform(this.FromDate, 'MM.dd.yyyy') + '-' + this.shared.datePipe.transform(this.ToDate, 'MM.dd.yyyy') : this.datetype(), 'Pace', 'Target', 'Variance', this.datetype() == 'C' ? this.shared.datePipe.transform(this.FromDate, 'MM.dd.yyyy') + '-' + this.shared.datePipe.transform(this.ToDate, 'MM.dd.yyyy') : this.datetype(), 'Pace', 'Target', 'Variance', this.datetype() == 'C' ? this.shared.datePipe.transform(this.FromDate, 'MM.dd.yyyy') + '-' + this.shared.datePipe.transform(this.ToDate, 'MM.dd.yyyy') : this.datetype(), 'Pace', 'Target', 'Variance', '% to Target']
    const headerRow = worksheet.addRow(Headings);
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
      cell.border = { right: { style: 'dotted' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });

    // Generate a unique file name
    this.exceldata.forEach((d: any) => {
      const Obj = [
        d['Rank'] == '' ? '-' : d['Rank'] == null ? '-' : d['Rank'],
        d['StoreName'] == '' ? '-' : d['StoreName'] == null ? '-' : d['StoreName'],
        d["New_MTD"] == '' ? '-' : d["New_MTD"] == null ? '-' : parseInt(d["New_MTD"]),
        d["New_Pace"] == '' ? '-' : d["New_Pace"] == null ? '-' : parseInt(d["New_Pace"]),
        d["New_Target"] == '' ? '-' : d["New_Target"] == null ? '-' : parseInt(d["New_Target"]),
        d["New_Variance"] == '' ? '-' : d["New_Variance"] == null ? '-' : parseInt(d["New_Variance"]),

        d["Used_MTD"] == '' ? '-' : d["Used_MTD"] == null ? '-' : parseInt(d["Used_MTD"]),
        d["Used_Pace"] == '' ? '-' : d["Used_Pace"] == null ? '-' : parseInt(d["Used_Pace"]),
        d["Used_Target"] == '' ? '-' : d["Used_Target"] == null ? '-' : parseInt(d["Used_Target"]),
        d["Used_Variance"] == '' ? '-' : d["Used_Variance"] == null ? '-' : parseInt(d["Used_Variance"]),

        d["Total_MTD"] == '' ? '-' : d["Total_MTD"] == null ? '-' : parseInt(d["Total_MTD"]),
        d["Total_Pace"] == '' ? '-' : d["Total_Pace"] == null ? '-' : parseInt(d["Total_Pace"]),
        d["Total_Target"] == '' ? '-' : d["Total_Target"] == null ? '-' : parseInt(d["Total_Target"]),
        d["Total_Variance"] == '' ? '-' : d["Total_Variance"] == null ? '-' : parseInt(d["Total_Variance"]),
        d["Total_Percentage"] == '' ? '-' : d["Total_Percentage"] == null ? '-' : parseInt(d["Total_Percentage"]) + '%',



      ];
      // console.log(Obj);
      const row = worksheet.addRow(Obj);
      row.font = { name: 'Arial', family: 4, size: 9 };
      row.eachCell((cell, number) => {
        cell.border = { right: { style: 'dotted' } };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        if (number >= 3 && number <= 14 && (this.filtertype == 'Gross')) {
          cell.numFmt = '$#,##0';
          cell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
        }
        if (number >= 3) {
          if (Obj[number] < 0) {
            row.getCell(number + 1).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } };
          }
        }

        // if (number == 11 || number == 12 || number == 13) {
        //   cell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
        // }
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
      row.worksheet.columns.forEach((column: any, columnIndex: any) => {
        let maxLength = 0;
        column.eachCell({ includeEmpty: true }, (cell: any) => {
          const length = cell.value ? cell.value.toString().length : 10;
          if (length > maxLength) {
            maxLength = length;
          }
        });
        column.width = maxLength < 10 ? 10 : maxLength + 2; // Set a minimum width of 10
      });
    });


    worksheet.getColumn(1).width = 20; worksheet.getColumn(1).alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
    worksheet.getColumn(2).width = 25;
    // worksheet.getColumn(3).width = 10;
    // worksheet.getColumn(4).width = 10;
    // worksheet.getColumn(5).width = 20;
    // worksheet.getColumn(6).width = 20;
    // worksheet.getColumn(7).width = 20;
    // worksheet.getColumn(8).width = 10;
    // worksheet.getColumn(9).width = 10;
    // worksheet.getColumn(10).width = 10;
    // worksheet.getColumn(11).width = 30;
    worksheet.addRow([]);
    workbook.xlsx.writeBuffer().then(buffer => {
      this.shared.exportToExcel(workbook, 'Sales Contest')
    });
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
