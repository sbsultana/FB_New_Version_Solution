import { Component, ViewChild, ElementRef, HostListener } from '@angular/core';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { Setdates } from '../../../../Core/Providers/SetDates/setdates';
import { BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { common } from '../../../../common';
import { Subscription } from 'rxjs';
import { environment } from '../../../../../environments/environment';
import { Stores } from '../../../../CommonFilters/stores/stores';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
import { CurrencyPipe } from '@angular/common';
@Component({
  selector: 'app-dashboard',
  imports: [SharedModule, BsDatepickerModule, Stores],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
})
export class Dashboard {
  TechData: any = [];
  NoData: boolean = false;
  otherstoreid: any = '';
  selectedotherstoreids: any = '';
  reportOpenSub!: Subscription;
  reportGetting!: Subscription;
  Pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;
  roleid: any;
  allordiff: any = "dr"
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
    const clickedInside = (event.target as HTMLElement).closest('.dropdown-toggle, .reportstores-card, .timeframe');
    if (!clickedInside) {
      this.activePopover = -1;
    }
  }
  constructor(
    public shared: Sharedservice, public setdates: Setdates, private comm: common, private cp: CurrencyPipe, private toast: ToastService,
  ) {
    this.shared.setTitle(this.comm.titleName + '-Tech Rate Exception');
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
    this.GetData();
  }
  AllorDiff(e: any) {
    this.allordiff = []
    const index = this.allordiff.findIndex((i: any) => i == e);
    if (index >= 0) {
      this.allordiff.splice(index, 1);
    } else {
      this.allordiff.push(e);
    }
  }
  setHeaderData() {
    const data = {
      title: 'Tech Rate Exception',
      stores: this.storeIds,
      datetype: 'MTD',
      groups: this.groupId,
      allordiff: this.allordiff,
    };
    this.shared.api.SetHeaderData({
      obj: data,
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
  isDesc: boolean = false;
  sortColumn: string = '';
  column: string = '';
  sortDirection: string = 'asc';
  sortData(property: any, flag?: any) {
    let direction: any = ''
    console.log(this.column)
    if (flag == undefined) {
      this.isDesc = !this.isDesc; //change the direction
      direction = this.isDesc ? 1 : -1;
    }
    else {
      direction = this.isDesc ? 1 : -1;
    }
    if (this.column == property) {
      this.sortDirection = this.isDesc ? 'asc' : 'desc';
      direction = this.isDesc ? 1 : -1;
    }
    else {
      this.sortDirection = 'asc'
      direction = 1
    }
    this.column = property;
    this.sortColumn = property;
  
    this.FilteredTechData.sort(function (a: any, b: any) {
      let valA = a[property];
      let valB = b[property];
      // Special case: parseInt for TechNumber
      if (property === 'TechNumber') {
        valA = parseInt(valA, 10);
        valB = parseInt(valB, 10);
      }
      if (valA < valB) {
        return -1 * direction;
      } else if (valA > valB) {
        return 1 * direction;
      } else {
        return 0;
      }
    });
   
  }
  FilteredTechData: any[] = [];
  filterTechData(): void {
    if (this.allordiff.toString() === 'all') {
      this.FilteredTechData = this.TechData;
    }
    else if (this.allordiff.toString() === 'exc') {
      // Exceptions > $11 with base condition
      this.FilteredTechData = this.TechData.filter(
        (item: { Diff: number; CDKRate: number; UKGRate: number }) =>
          item.UKGRate !== 0 &&
          item.CDKRate !== 0 &&
          item.Diff !== 0 &&
          item.Diff > 11
      );
    }
    else if (this.allordiff.toString() === 'cdkukg') {
      // CDK > UKG (negative Diff) with base condition
      this.FilteredTechData = this.TechData.filter(
        (item: { Diff: number; CDKRate: number; UKGRate: number }) =>
          item.UKGRate !== 0 &&
          item.CDKRate !== 0 &&
          item.Diff !== 0 &&
          item.Diff < 0
      );
    }
    else {
      // Default differences (exclude > $11)
      this.FilteredTechData = this.TechData.filter(
        (item: { Diff: number; CDKRate: number; UKGRate: number }) =>
          item.UKGRate !== 0 &&
          item.CDKRate !== 0 &&
          item.Diff !== 0 &&
          !(item.Diff > 11)
      );
    }
    if (this.FilteredTechData.length == 0) {
      this.NoData = true
    }
    else {
      this.NoData = false
    }
    console.log(this.FilteredTechData.length)
  }
  GetData() {
    this.TechData = [];
    this.FilteredTechData = [];
    this.shared.spinner.show();
    const obj = {
      StoreID: this.selectedotherstoreids != undefined && this.selectedotherstoreids != '' && this.selectedotherstoreids != null ?
        (this.storeIds != '' ? this.storeIds + ',' + this.selectedotherstoreids.toString() : this.selectedotherstoreids.toString()) : this.storeIds,
    };
    const curl = environment.apiUrl + this.comm.routeEndpoint + 'GetTechRateException';
    this.shared.api
      .postmethod(this.comm.routeEndpoint + 'GetTechRateException', obj)
      .subscribe(
        (res) => {
          const currentTitle = document.title;
          this.shared.api.logSaving(curl, {}, '', res.message, currentTitle);
          if (res.status == 200) {
            if (res.response != undefined) {
              if (res.response.length > 0) {
                this.TechData = res.response;
                this.filterTechData();
                this.shared.spinner.hide()

              } else {
                this.shared.spinner.hide();
                this.NoData = true;
              }
            } else {
              this.shared.spinner.hide();
              this.NoData = true;
            }
          } else {
            this.toast.show(res.status,'danger','Error');
            this.shared.spinner.hide();
            this.NoData = true;
            // this.spinnerLoader=false;
          }
        },
        (error) => {
          // this.toast.error('502 Bad Gate Way Error', '');
          this.shared.spinner.hide();
          this.NoData = true;
        }
      );
  }
  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    }
    return false;
  }
  expandorcollapse(ind: any, e: any, ref: any, Item: any) {
    let id = (e.target as Element).id;
    if (id == 'D_' + ind) {
      if (ref == '-') {
        Item.Dealerx = '+';
      }
      if (ref == '+') {
        Item.Dealerx = '-';
      }
    }
  }
  SPRstate: any;
  ngAfterViewInit(): void {
      this.shared.api.getStores().subscribe((res: any) => {
      if (this.comm.pageName == 'Tech Rate Exception') {
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
        this.SPRstate = res.obj.state;
        if (res.obj.title == 'Tech Rate Exception') {
          if (res.obj.state == true) {
            this.exportToExcel();
          }
        }
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res) => {
      if (this.print != undefined) {
        if (res.obj.title == 'Tech Rate Exception') {
          if (res.obj.statePrint == true) {
          //  this.GetPrintData();
          }
        }
      }
    });
    this.Pdf = this.shared.api.getExportToPDFAllReports().subscribe((res) => {
      if (this.Pdf != undefined) {
        if (res.obj.title == 'Tech Rate Exception') {
          if (res.obj.statePDF == true) {
         //   this.generatePDF();
          }
        }
      }
    });
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res) => {
      if (this.email != undefined) {
        if (res.obj.title == 'Tech Rate Exception') {
          if (res.obj.stateEmailPdf == true) {
         //   this.sendEmailData(res.obj.Email, res.obj.notes, res.obj.from);
          }
        }
      }
    });
  }
    activePopover: number = -1;

  togglePopover(popoverIndex: number) {
    this.activePopover = this.activePopover === popoverIndex ? -1 : popoverIndex;
  }

    viewreport() {
    this.activePopover = -1
    if (this.storeIds.length == 0 ) {
      this.toast.show('Please select atleast one Store', 'warning','Warning');
    }
   
    else {
    this.GetData()
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
  ExcelStoreNames: any = [];
  exportToExcel() {
    let storeNames: any = [];
    let store = this.storeIds.split(',');
    // const obj = {
    //   id: this.groups,
    //   userid: JSON.parse(localStorage.getItem('UserDetails')!).userid,
    // };
    // this.shared.api
    //   .postmethodOne(this.comm.routeEndpoint+'GetStoresbyGroupuserid', obj)
    //   .subscribe((res: any) => {
    storeNames = this.comm.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores.filter((item: any) =>
      store.some((cat: any) => cat === item.ID.toString())
    );
    // console.log(store,res.response.length);
    if (store.length == this.comm.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores.length) {
      this.ExcelStoreNames = 'All Stores'
    } else {
      this.ExcelStoreNames = storeNames.map(function (a: any) {
        return a.storename;
      });
    }
    const SalespersonData = this.FilteredTechData.map(
      (_arrayElement: any) => Object.assign({}, _arrayElement)
    );
    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('Tech Rate Exception');
    worksheet.views = [
      {
        state: 'frozen',
        ySplit: 11, // Number of rows to freeze (2 means the first two rows are frozen)
        topLeftCell: 'A12', // Specify the cell to start freezing from (in this case, the third row)
        showGridLines: false,
      },
    ];
    worksheet.addRow('');
    const titleRow = worksheet.addRow(['Tech Rate Exception']);
    titleRow.eachCell((cell, number) => {
      cell.alignment = { indent: 1, vertical: 'middle', horizontal: 'left' };
    });
    titleRow.font = { name: 'Arial', family: 4, size: 12, bold: true };
    titleRow.worksheet.mergeCells('A2', 'D2');
    worksheet.addRow('');
    // const DateToday = this.shared.datePipe.transform(
    //   new Date(),
    //   'MM/dd/yyyy h:mm:ss a'
    // );
    const DATE_EXTENSION = this.shared.datePipe.transform(
      new Date(),
      'MMddyyyy'
    );
    // worksheet.addRow([DateToday]).font = { name: 'Arial', family: 4, size: 9 };
    // const ReportFilter = worksheet.addRow(['Report Controls :']);
    // ReportFilter.font = { name: 'Arial', family: 4, size: 10, bold: true };
    // const Timeframe = worksheet.addRow(['Timeframe :']);
    // Timeframe.getCell(1).font = {
    //   name: 'Arial',
    //   family: 4,
    //   size: 9,
    //   bold: true,
    // };
    // const timeframe = worksheet.getCell('B6');
    // timeframe.value = this.FromDate + ' to ' + this.ToDate;
    // timeframe.font = { name: 'Arial', family: 4, size: 9 };
    const Groups = worksheet.getCell('A7');
    Groups.value = 'Group :';
    Groups.font = { name: 'Arial', family: 4, size: 9, bold: true };
    const groups = worksheet.getCell('B7');
    groups.value =
      this.comm.groupsandstores.filter((val: any) => val.sg_id == this.groupId.toString())[0].sg_name;
    groups.font = { name: 'Arial', family: 4, size: 9 };
    groups.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
    worksheet.mergeCells('B8', 'K10');
    const Stores = worksheet.getCell('A8');
    Stores.value = 'Stores :'
    const stores = worksheet.getCell('B8');
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
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber > 3 && rowNumber < 11) {
        // Apply styles, formatting, or other actions for even rows
        row.getCell(1).alignment = {
          vertical: 'middle', horizontal: 'left', indent: 1
        }
      }
    });
    // worksheet.mergeCells('A12', 'C12');
    // let MTD = worksheet.getCell('A12');
    // MTD.value = 'MTD';
    // MTD.alignment = { vertical: 'middle', horizontal: 'center' };
    // MTD.font = {
    //   name: 'Arial',
    //   family: 4,
    //   size: 9,
    //   bold: true,
    //   color: { argb: 'FFFFFF' },
    // };
    // MTD.fill = {
    //   type: 'pattern',
    //   pattern: 'solid',
    //   fgColor: { argb: '2a91f0' },
    //   bgColor: { argb: 'FF0000FF' },
    // };
    // MTD.border = { right: { style: 'thin' } };
    // worksheet.mergeCells('D12', 'H12');
    // let UnitCredit = worksheet.getCell('D12');
    // UnitCredit.value = 'Unit Credit';
    // UnitCredit.alignment = { vertical: 'middle', horizontal: 'center' };
    // UnitCredit.font = {
    //   name: 'Arial',
    //   family: 4,
    //   size: 9,
    //   bold: true,
    //   color: { argb: 'FFFFFF' },
    // };
    // UnitCredit.fill = {
    //   type: 'pattern',
    //   pattern: 'solid',
    //   fgColor: { argb: '2a91f0' },
    //   bgColor: { argb: 'FF0000FF' },
    // };
    // UnitCredit.border = { right: { style: 'thin' } };
    // worksheet.mergeCells('I12', 'L12');
    // let FrontGross = worksheet.getCell('I12');
    // FrontGross.value = 'Front Gross';
    // FrontGross.alignment = { vertical: 'middle', horizontal: 'center' };
    // FrontGross.font = {
    //   name: 'Arial',
    //   family: 4,
    //   size: 9,
    //   bold: true,
    //   color: { argb: 'FFFFFF' },
    // };
    // FrontGross.fill = {
    //   type: 'pattern',
    //   pattern: 'solid',
    //   fgColor: { argb: '2a91f0' },
    //   bgColor: { argb: 'FF0000FF' },
    // };
    // FrontGross.border = { right: { style: 'thin' } };
    let Headings = [
      'Tech Name',
      'Tech Number',
      'CDK Rate',
      'UKG Rate',
      'Diff',
    ];
    const headerRow = worksheet.addRow(Headings);
    headerRow.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: false,
      color: { argb: 'FFFFFF' },
    };
    headerRow.alignment = { indent: 1, vertical: 'middle', horizontal: 'center' };
    headerRow.eachCell((cell, number) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '788494' },
        bgColor: { argb: 'FF0000FF' },
      };
      cell.border = { right: { style: 'thin' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });
    for (const d of SalespersonData) {
      const Data = worksheet.addRow([
        d.TechName == '' ? '-' : d.TechName == null ? '-' : d.TechName,
        d.TechNumber == '' ? '-' : d.TechNumber == null ? '-' : d.TechNumber,
        d.CDKRate == '' ? '-' : d.CDKRate == null ? '-' : d.CDKRate,
        d.UKGRate == '' ? '-' : d.UKGRate == null ? '-' : d.UKGRate,
        d.Diff == '' ? '-' : d.Diff == null ? '-' : d.Diff
      ]);
      // Data1.outlineLevel = 1; // Grouping level 1
      Data.font = { name: 'Arial', family: 4, size: 8 };
      Data.alignment = { vertical: 'middle', horizontal: 'center' };
      Data.getCell(2).alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
      Data.getCell(3).alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
      Data.eachCell((cell, number) => {
        cell.border = { right: { style: 'thin' } };
        if (number > 3 && number < 9) {
          cell.numFmt = '#,##0.0';
        }
        if (number > 8 && number < 13) {
          cell.numFmt = '$#,##0';
        }
      });
      if (Data.number % 2) {
        Data.eachCell((cell, number) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'e5e5e5' },
            bgColor: { argb: 'FF0000FF' },
          };
        });
      }
      if (d.StoreName == null) {
        Data.eachCell((cell) => {
          cell.font = { name: 'Arial', family: 4, size: 9, bold: true };
          cell.border = {
            top: { style: 'thin' },
            bottom: { style: 'thin' },
          };
        });
      }
    }
    worksheet.getColumn(1).width = 15;
    worksheet.getColumn(2).width = 30;
    worksheet.getColumn(3).width = 30;
    worksheet.getColumn(4).width = 15;
    worksheet.getColumn(5).width = 15;
    worksheet.getColumn(6).width = 15;
    worksheet.getColumn(7).width = 15;
    worksheet.getColumn(8).width = 15;
    worksheet.getColumn(9).width = 15;
    worksheet.getColumn(10).width = 15;
    worksheet.getColumn(11).width = 15;
    worksheet.getColumn(12).width = 15;
    worksheet.addRow([]);
    workbook.xlsx.writeBuffer().then((data: any) => {
      const blob = new Blob([data], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
     this.shared.exportToExcel(workbook, 'Tech Rate Exception_' + DATE_EXTENSION );
    });
    // });
  }
}