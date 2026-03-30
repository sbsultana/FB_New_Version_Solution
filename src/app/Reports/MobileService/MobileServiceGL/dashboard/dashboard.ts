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
  imports: [SharedModule, BsDatepickerModule, DateRangePicker],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
})
export class Dashboard {
  ExpensesBlock: any[] = []
  ExpensesBlockNoData: any = ''


  NoData: boolean = false;
  responcestatus: any = '';
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
  groups: any = 0;
  storeIds: any = 0;



  FromDate: any = '';
  ToDate: any = '';
  minDate!: Date;
  maxDate!: Date;
  DateType: any = 'MTD';
  displaytime: any = '';
  DupFromDate: any = '';
  DupToDate: any = ''

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
    this.shared.setTitle(this.comm.titleName + '-Mobile Service GL');
      this.initializeDates('MTD')

      if (this.comm.MobileServiceGL != undefined && this.comm.MobileServiceGL.length > 0) {
        this.stores = this.comm.MobileServiceGL;
        if (this.storeIds == undefined || this.storeIds.length == 0) {
          this.getStoresandGroupsValues('FL')
        }
      }
    this.setHeaderData();
  }
  asoftime: any = [];
  ngOnInit() { }
  initializeDates(type: any) {
    let dates: any = this.setdates.setDates(type)
    this.FromDate = dates[0];
    this.ToDate = dates[1];
    localStorage.setItem('time', type);
    this.DateType = type
    this.setDates(this.DateType)

  }
  setHeaderData() {
    // alert('HI')
    const data = {
      title: 'Mobile Service GL',
      stores: this.storeIds,
      fromdate: this.FromDate,
      todate: this.ToDate,
      groups: this.groups,
    };
    this.shared.api.SetHeaderData({
      obj: data,
    });
    // this.getServiceData()
  }
  getServiceData() {
    this.ExpensesBlockNoData = ''
    this.ExpensesBlock = []
    this.responcestatus = '';
    this.GetData('')
  }
  incomeGrossTotal: any = 0;
  GetData(block: any) {
      this.DupFromDate = this.FromDate;
    this.DupToDate = this.ToDate
    const obj = {
      Startdate: this.FromDate.replaceAll('/', '-'),
      Enddate: this.ToDate.replaceAll('/', '-'),
      StoreID: this.storeIds.toString(),
    };
    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetMobileServiceGL', obj).subscribe(
      (res) => {
        console.log(res.response, 'Response');
        if (res && res.response && res.response.length > 0) {
          if (block == '') {
            let Expense = res.response
            Expense.some(function (x: any) {
              if (x.Accounts != undefined) {
                x.Accounts = JSON.parse(x.Accounts);
              }
            });
            let data = Expense.reduce(
              (r: any, { ASG_TYPE }: any) => {
                if (!r.some((o: any) => o.ASG_TYPE == ASG_TYPE)) {
                  r.push({
                    ASG_TYPE,
                    seq: ASG_TYPE == 'Income' ? 1 : (ASG_TYPE == 'Selling Expenses' ? 2 : (ASG_TYPE == 'Mobile Dept Selling Gross' ? 3 : (ASG_TYPE == 'Add Back' ? 4 : 5))),
                    // seq: ASG_TYPE == 'Add Back' ? 3 : (ASG_TYPE == 'Selling Expenses' ? 2 : (ASG_TYPE == 'Adj Mobile Service Profit (After Offset)' ? 4 : 1)),
                    subdata: Expense.filter(
                      (v: any) => v.ASG_TYPE == ASG_TYPE
                    ).sort((a: any, b: any) => a.acct_seq - b.acct_seq),
                  });
                }
                return r;
              },
              []
            ).sort((a: any, b: any) => a.seq - b.seq);
            this.ExpensesBlock = data;
            let incomeblock: any = []
            this.ExpensesBlock && this.ExpensesBlock.length > 0 ? incomeblock = this.ExpensesBlock.filter((val: any) => val.ASG_TYPE == 'Income')[0]?.subdata.filter((e: any) => e.ASG_Department == 'Total')[0]?.Gross : ''
            this.incomeGrossTotal = incomeblock
            console.log(incomeblock, this.incomeGrossTotal, 'Income Bloc');
            console.log(this.ExpensesBlock, 'Expenses Block');
          }
        }
        else if (res.status == 200) {
          this.emptyData(block)
        } else {
          this.toast.show(res.status, 'danger', 'Error');
          this.shared.spinner.hide();
          this.emptyData(block)
        }
      },
      (error) => {
        this.toast.show('502 Bad Gate Way Error', 'danger', 'Error');
        this.shared.spinner.hide();
        this.emptyData(block)
      }
    );
  }
  emptyData(block: any) {
    this.ExpensesBlock ? (this.ExpensesBlock.length > 0 ? this.ExpensesBlockNoData = '' : this.ExpensesBlockNoData = 'No Data Found!!') : this.ExpensesBlockNoData = 'No Data Found!!'
  }
  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    } else if (value < 0) {
      return false;
    }
    return true;
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
  ngAfterViewInit() {
    // this.shared.api.getStores().subscribe((res: any) => {
    //   if (this.comm.pageName == 'Mobile Service GL') {
    //     if (res.obj.storesData != undefined) {
    //       this.groupsArray = res.obj.storesData;
    //       this.getGroupBaseStores(this.groups, 'FL')
    //     }
    //   }
    // })
    this.shared.api.getMobileService().subscribe((res: any) => {
      if (this.comm.pageName == 'Mobile Service GL') {
        if (res.obj.storesData != undefined) {
          this.stores = res.obj.storesData;
          this.storeIds = []
          if (this.storeIds == undefined || this.storeIds.length == 0) {
            this.getGroupBaseStores(this.groups, 'FL')
            console.log(this.stores, 'S');
          }
        }
      }
    })
    this.shared.api.getAllStores().subscribe((res: any) => {
      if (this.comm.pageName == 'Mobile Service GL') {
        if (res.obj.storesData != undefined) {
          this.groupsArray = res.obj.storesData;
          if (this.storeIds == undefined || this.storeIds.length == 0) {
            this.getGroupBaseStores(this.groups, 'FL')
          }
        }
      }
    })
    this.excel = this.shared.api.getExportToExcelAllReports().subscribe((res) => {
      if (this.excel != undefined) {
        if (res.obj.title == 'Mobile Service GL') {
          if (res.obj.state == true) {
            this.exportToExcel();
          }
        }
      }
    });
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res) => {
      if (this.email != undefined) {
        if (res.obj.title == 'Mobile Service GL') {
          if (res.obj.stateEmailPdf == true) {
            // this.sendEmailData(res.obj.Email, res.obj.notes, res.obj.from);
          }
        }
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res) => {
      if (this.print != undefined) {
        if (res.obj.title == 'Mobile Service GL') {
          if (res.obj.statePrint == true) {
            // this.GetPrintData();
          }
        }
      }
    });
    this.Pdf = this.shared.api.getExportToPDFAllReports().subscribe((res) => {
      if (this.Pdf != undefined) {
        if (res.obj.title == 'Mobile Service GL') {
          if (res.obj.statePDF == true) {
            // this.generatePDF();
          }
        }
      }
    });
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
  individualgroups(e: any) {
    this.groups = []
    this.groups.push(e.sg_id);
    // this.groupName = this.groupsArray.filter((val: any) => val.sg_id == this.groups[0])[0].sg_name;
    this.groupName = 'Recon Centers';
    this.getGroupBaseStores(this.groups.toString(), 'FG')
  }
  spinnerLoader: boolean = false;
  getGroupBaseStores(id: any, block: any) {
    this.spinnerLoader = true
    this.stores = this.comm.MobileServiceGL && this.comm.MobileServiceGL.length > 0 ? this.comm.MobileServiceGL : []
    this.spinnerLoader = false
    this.storeIds = []
    if (block == 'FG') {
      this.storeIds = this.stores.map(function (a: any) {
        return a.ac_as_id;
      });
      this.groupName = 'Recon Centers';
    }
    if (block == 'FL') {
      this.getStoresandGroupsValues('FL')
    }
  }
  getStoresandGroupsValues(type?: any) {
    // alert('Hello')
    let data = JSON.parse(localStorage.getItem('UserDetails')!)
    let dupStores = this.stores.map(function (a: any) {
      return a.ac_as_id;
    });
    let duptokenstores: any = [];
    type == 'FL' ? JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids.indexOf(',') > 0 ? duptokenstores = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids.split(',') : duptokenstores.push(JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids) : ''
    const intersection = dupStores.filter((element: any) => duptokenstores.includes(element.toString()));
    console.log(intersection, duptokenstores, dupStores);
    this.storeIds = []
    type == 'FL' ? this.storeIds = intersection :
      this.storeIds = this.stores.map(function (a: any) {
        return a.ac_as_id;
      });
    console.log(this.storeIds, '.............');
    this.setHeaderData()
    if (this.storeIds && this.storeIds.length > 0 && this.FromDate != '' && this.ToDate != '') {
      this.getServiceData()
    }
    if (intersection.length == 0) {
      this.ExpensesBlockNoData = 'Please Select Store'
      this.ExpensesBlock = []
      this.responcestatus = '';
    }
    if (this.stores.length == this.storeIds.length) {
      // this.groupName = this.groupsArray.filter((val: any) => val.sg_id == this.groups[0])[0].sg_name;
      this.groupName = 'Recon Centers';
    }
    if (this.storeIds.length == 1) {
      this.storename = this.stores.filter((val: any) => val.ac_as_id == this.storeIds.toString())[0].storename
    }
  }
  individualStores(e: any) {
    const index = this.storeIds.findIndex((i: any) => i == e.ac_as_id);
    if (index >= 0) {
      this.storeIds.splice(index, 1);
    } else {
      this.storeIds.push(e.ac_as_id);
    }
    if (this.storeIds.length == 1) {
      this.storename = this.stores.filter((val: any) => val.ac_as_id == this.storeIds.toString())[0].storename
    }
  }
  // allstores() {
  //   if (this.storeIds.length == this.stores.length) {
  //     this.storeIds = [];
  //   } else {
  //     this.storeIds = this.stores.map(function (a: any) {
  //       return a.ac_as_id;
  //     });
  //     // this.groupName = this.groupsArray.filter((val: any) => val.sg_id == this.groups[0])[0].sg_name;
  //     this.groupName = 'Recon Centers';
  //   }
  // }
  allstores(state: any) {
    if (state == 'N') {
      this.storeIds = [];
    } else if (state == 'Y') {
      this.storeIds = this.stores.map(function (a: any) {
        return a.ac_as_id;
      });
      this.groupName = 'Recon Centers';
    }

  }
  viewreport() {
    this.activePopover = -1
    if (this.storeIds.length == 0) {
      this.toast.show('Please Select Atleast One Store', 'warning', 'Warning');
    }
    else {
      // const data = {
      //   Reference: 'Mobile Service GL',
      //   storeValues: this.storeIds.toString(),
      //   groups: this.groups.toString(),
      //   selectedrecon: this.selectedRecon
      // };
      // this.shared.api.SetReports({
      //   obj: data,
      // });
      this.setHeaderData();
      this.getServiceData()
    }
  }

  viewDeal(dealData: any) {
    // const modalRef = this.ngbmodal.open(DealRecapComponent, { size: 'md', windowClass: 'connectedmodal' });
    // modalRef.componentInstance.data = { dealno: dealData.dealno, storeid: dealData.dealerid, stock: dealData.Stock, vin: dealData.VIN, custno: dealData?.ad_custid }; // Pass data to the modal component    
    // modalRef.result.then((result) => {
    //   console.log(result); // Handle modal close result
    // }, (reason) => {
    //   console.log(`Dismissed: ${reason}`); // Handle dismiss reason
    // });
  }
  getGP(data: any) {
    let val = (data / this.incomeGrossTotal) * -100
    // if (val) {
    //   return val
    // } else {
    //   return
    // }
    if (!isNaN(val)) {
      return Number(Math.abs(val).toFixed(1)); // always positive
    } else {
      return
    }
  }

  ExcelStoreNames: any = []
  exportToExcel() {
    let storeNames: any[] = [];
    const store = this.storeIds
    storeNames = this.comm.MobileServiceGL.filter((item: any) =>
      store.includes(item.ac_as_id)
    );
    if (store.length == this.comm.MobileServiceGL.length) {
      this.ExcelStoreNames = 'All Stores'
    } else {
      this.ExcelStoreNames = storeNames.map(function (a: any) {
        return a.storename;
      });
    }
    console.log(store, this.ExcelStoreNames, storeNames);
    // Setup Excel
    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('Mobile Service GL');
    const DATE_EXTENSION = this.shared.datePipe.transform(new Date(), 'MMddyyyy');
    worksheet.views = [{ state: 'frozen', ySplit: 5, topLeftCell: 'A6', showGridLines: false }];
    // Header section (above grid)
    worksheet.addRow([]);
    const titleRow = worksheet.addRow(['Mobile Service GL']);
    titleRow.font = { bold: true, size: 12 };
    worksheet.addRow([]);
    const Stores1 = worksheet.getCell('A3');
    Stores1.value = 'Stores :';
    worksheet.mergeCells('B3', 'Z3');
    const stores1 = worksheet.getCell('B3');
    stores1.value = this.ExcelStoreNames.toString().replaceAll(',', ', ');
    stores1.font = { name: 'Arial', family: 4, size: 9 };
    stores1.alignment = { vertical: 'top', horizontal: 'left', wrapText: true, };
    const Timeframe = worksheet.getCell('A4');
    Timeframe.value = 'Time Frame :';
    worksheet.mergeCells('B4', 'Z4');
    const timeframe = worksheet.getCell('B4');
    timeframe.value = this.shared.datePipe.transform(this.FromDate, 'MM/dd/yyyy') + ' - ' + this.shared.datePipe.transform(this.ToDate, 'MM/dd/yyyy');
    timeframe.font = { name: 'Arial', family: 4, size: 9 };
    timeframe.alignment = { vertical: 'top', horizontal: 'left', wrapText: true, };
    const columnWidths = new Array(6).fill(10);
    // Main Data Rows
    this.ExpensesBlock.forEach((item: any, index: number) => {
      if (index < 2) {
        const headers = [item.ASG_TYPE, 'ACCT #', 'Name/Desc', 'Sales', 'Gross', 'GP%'];
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
        const columnWidths = new Array(headers.length).fill(10);
      }
      item.subdata.forEach((sub: any, j: any) => {
        const subData = [
          sub.ASG_Department != 'Total' ? sub.ASG_Department : '',
          sub.ASG_Department != 'Total' ? sub.act : '', sub.ASG_NAME, sub.saleamt, sub.Gross,
          item.ASG_TYPE == 'Income' ? (sub.GP ? sub.GP + '%' : '') : (this.getGP(sub.Gross) ? this.getGP(sub.Gross) + '%' : '')
        ];
        const subRow = worksheet.addRow(subData);
        subRow.font = { size: 9 };
        if (typeof sub.saleamt === 'number' || typeof sub.Gross === 'number') {
          subRow.getCell(4).numFmt = '$#,##0.00';
          subRow.getCell(5).numFmt = '$#,##0.00';
        }
        subRow.getCell(4).alignment = { horizontal: 'right', indent: 1 };
        subRow.getCell(5).alignment = { horizontal: 'right', indent: 1 };
        subRow.getCell(6).alignment = { horizontal: 'right', indent: 1 };
        subRow.eachCell((cell) => {
          // cell.fill = {
          //   type: 'pattern',
          //   pattern: 'solid',
          //   fgColor: { argb: 'd2deed' },
          // };
          cell.alignment = { ...cell.alignment, indent: 1 };
        });
        subData.forEach((val, i) => {
          const length = val?.toString().length || 0;
          columnWidths[i] = Math.max(columnWidths[i], length);
        });
      });
      columnWidths.forEach((width, i) => {
        worksheet.getColumn(i + 1).width = Math.max(15, width + 2);
      });
    })

    this.shared.exportToExcel(workbook, `Mobile Service GL_${DATE_EXTENSION}.xlsx`);

  }
}