import { Component, Injector, HostListener } from '@angular/core';
import { Api } from '../../../../Core/Providers/Api/api';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { NgbActiveModal, NgbModal } from '@ng-bootstrap/ng-bootstrap';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
import { forkJoin, Subscription } from 'rxjs';
import { BsDatepickerConfig, BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { Stores } from '../../../../CommonFilters/stores/stores';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { Workbook } from 'exceljs';
import * as FileSaver from 'file-saver';
import { DatePipe } from '@angular/common';
const EXCEL_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
const EXCEL_EXTENSION = '.xlsx';

@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [SharedModule, BsDatepickerModule, Stores],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
})


export class Dashboard {
  UnappliedData: any = [];
  keys: string[] = [];
  pastThreeMonths: Date[] = [];
  expanddataarray: any = [];
  NoData: any = ''
  reporttotals: any = 'B'
  stores: any = []
  groupsArray: any = [];
  storename: any = ''
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  groupId: any = 0;
  storeIds: any = 0;
  TotalSortPosition: any = 'B';
  excel!: Subscription;

  storesFilterData: any = {
    'groupsArray': this.groupsArray, 'groupId': this.groupId, 'storesArray': this.stores, 'storeids': '1', 'type': 'M', 'others': 'N',
    'groupName': this.groupName, 'storename': this.storename, storecount: null, 'storedisplayname': this.storedisplayname
  };
  activePopover: number = -1;

  month!: Date;
  DupMonth: Date = new Date();
  DuplicatDate!: Date;
  minDate!: Date;
  maxDate!: Date;
  bsConfig: Partial<BsDatepickerConfig> = {
    dateInputFormat: 'MMMM/YYYY',
    minMode: 'month'
  };
  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent) {
    const clickedInside = (event.target as HTMLElement).closest('.dropdown-toggle, .reportstores-card , .timeframe');
    if (!clickedInside) {
      this.activePopover = -1;
    }
  }
  constructor(public apiSrvc: Api, private ngbmodalActive: NgbActiveModal, private datepipe: DatePipe,
    private toast: ToastService, private injector: Injector, public shared: Sharedservice,) {

    let today = new Date();

    // current month (no subtraction)
    this.month = new Date(today.getFullYear(), today.getMonth(), 1);

    // maxDate → current month
    this.maxDate = new Date(today.getFullYear(), today.getMonth(), 1);

    // minDate → 3 years back
    this.minDate = new Date();
    this.minDate.setFullYear(today.getFullYear() - 3);
    this.maxDate = new Date();
    this.minDate = new Date();
    this.minDate.setFullYear(this.maxDate.getFullYear() - 3);
    this.maxDate.setMonth(this.maxDate.getMonth() - 1);
    if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
      // this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.ustores.split(',')
      this.groupId = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Preferences
      this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids.split(',')
    }
    console.log('store displayname', this.storedisplayname);
    if (this.shared.common.groupsandstores.length > 0) {
      this.groupsArray = this.shared.common.groupsandstores.filter((val: any) => val.sg_id != this.shared.common.reconID);
      this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
      this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_Name : this.groupName = ''
      this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
      this.getStoresandGroupsValues()
    }
    this.shared.setTitle(this.shared.common.titleName + '-UnappliedTimeReport')

    this.setHeaderData()
    this.GetUnappliedReport('', 0, '', '');
  }

  //getstores
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

  //dates
  onOpenCalendar(container: any) {
    container.setViewMode('month');
    container.monthSelectHandler = (event: any): void => {
      container.value = event.date;
      this.month = event.date;
      return;
    };
  }
  changeDate(e: any) {
    console.log(e);
    this.month = e;
  }
  getPastThreeMonths(): void {
    this.pastThreeMonths = [];

    const currentDate = new Date(this.month);

    for (let i = 0; i < 3; i++) {
      const tempDate = new Date(currentDate);
      tempDate.setMonth(currentDate.getMonth() - i);
      tempDate.setDate(1);

      this.pastThreeMonths.push(tempDate);
    }
  }
  setHeaderData() {
    const data = {
      title: 'Unapplied Time Report',
      stores: this.storeIds,
      groups: this.groupId,
      month: this.month
    };
    this.apiSrvc.SetHeaderData({
      obj: data,
    });
  }

  isDesc: boolean = false;
  column: string = '';

  sort(property: string) {

    this.isDesc = !this.isDesc;
    const direction = this.isDesc ? 1 : -1;


    const totalRow = this.UnappliedData.find(
      (x: any) => x.Store_Name?.toUpperCase() === 'REPORTS TOTAL'
    );

    let storeRows = this.UnappliedData.filter(
      (x: any) => x.Store_Name?.toUpperCase() !== 'REPORTS TOTAL'
    );

    storeRows.sort((a: any, b: any) => {
      const valA = a[property] ?? '';
      const valB = b[property] ?? '';

      return valA.toString().localeCompare(valB.toString()) * direction;
    });

    this.UnappliedData = totalRow
      ? [...storeRows, totalRow]
      : storeRows;
  }

  //excel
  ExcelStoreNames: any = [];

  exportToExcel() {

    //  STORE LOGIC
    const storeIdsArr = Array.isArray(this.storeIds)
      ? this.storeIds.map((x: any) => x.toString())
      : this.storeIds.toString().split(',');

    const groupData = this.shared.common.groupsandstores
      .find((v: any) => v.sg_id == this.groupId);

    if (!groupData) {
      this.toast.show('Group not found', 'warning', 'Warning');
      return;
    }

    const allStoreIds = groupData.Stores.map((s: any) => s.ID.toString());

    const selectedStores = groupData.Stores.filter((item: any) =>
      storeIdsArr.includes(item.ID.toString())
    );

    const isAllSelected =
      storeIdsArr.length === allStoreIds.length &&
      storeIdsArr.every((id: any) => allStoreIds.includes(id));

    this.ExcelStoreNames = isAllSelected
      ? 'All Stores'
      : selectedStores.map((a: any) => a.storename);

    const SalesData = this.UnappliedData || [];

    if (!SalesData.length) {
      this.toast.show('No data available', 'warning', 'Warning');
      return;
    }

    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet('Unapplied Time Report');

    worksheet.views = [{ showGridLines: false }];

    worksheet.addRow([]);

    // TITLE
    const titleRow = worksheet.addRow(['Unapplied Time Report']);
    titleRow.font = { name: 'Arial', size: 12, bold: true };
    titleRow.alignment = { horizontal: 'left', indent: 1 };
    worksheet.mergeCells('A2', 'D2');

    worksheet.addRow([]);

    const DateToday = this.datepipe.transform(new Date(), 'MM/dd/yyyy h:mm:ss a');
    const DATE_EXTENSION = this.datepipe.transform(new Date(), 'MMddyyyy');

    worksheet.addRow([DateToday]).font = { size: 9 };

    worksheet.addRow(['Report Controls :']).font = { bold: true };

    //  GROUP
    worksheet.getCell('B8').value = 'Group :';
    worksheet.getCell('D8').value = groupData.sg_name;

    // STORES
    worksheet.getCell('B9').value = 'Stores :';
    worksheet.mergeCells('D9', 'O11');

    const storesCell = worksheet.getCell('D9');

    storesCell.value = Array.isArray(this.ExcelStoreNames)
      ? this.ExcelStoreNames.join(', ')
      : this.ExcelStoreNames;

    storesCell.alignment = {
      wrapText: true,
      vertical: 'top',
      horizontal: 'left'
    };

    worksheet.addRow([]);

    // MONTH HEADER (WORK MIX STYLE)
    let start = 2;

    this.pastThreeMonths.forEach((val: any) => {

      worksheet.mergeCells(18, start, 18, start + 3);

      const cell = worksheet.getCell(18, start);

      cell.value = this.datepipe.transform(val, 'MMM yyyy');

      cell.alignment = { horizontal: 'center', vertical: 'middle' };

      cell.font = {
        bold: true,
        color: { argb: 'FFFFFF' }
      };

      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '2a91f0' }
      };

      //  THIN FULL BORDER (FIXED)
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };

      start += 4;
    });

    //  HEADERS (WORK MIX STYLE)
    let Heads: any = ['Store Name'];

    this.keys.forEach(() => {
      Heads.push('Unapplied Time', '% of Labor Sales', 'Budget', 'Diff');
    });

    const headerRow = worksheet.addRow(Heads);

    headerRow.font = {
      name: 'Arial',
      size: 9,
      bold: true,
      color: { argb: 'FFFFFF' }
    };

    headerRow.alignment = {
      vertical: 'middle',
      horizontal: 'center'
    };

    headerRow.eachCell((cell: any) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '2a91f0' }
      };

      //  FULL GRID (NOT DOTTED)
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
    });

    // DATA
    SalesData.forEach((d: any) => {

      let row: any = [d.Store_Name || '-'];

      this.keys.forEach((val: any) => {
        row.push(
          d[val]?.Mtd ?? '-',
          d[val]?.Labor_Sale_Per ? d[val].Labor_Sale_Per + '%' : '-',
          d[val]?.Budget ?? '-',
          d[val]?.Diff ?? '-'
        );
      });

      const excelRow = worksheet.addRow(row);

      excelRow.eachCell((cell: any, colNumber: number) => {

        //  FORMAT
        if (![3, 7, 11].includes(colNumber)) {
          cell.numFmt = '$#,##0';
        }

        //  FULL GRID
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };

        cell.alignment = {
          vertical: 'middle',
          horizontal: colNumber === 1 ? 'left' : 'center'
        };

        //  NEGATIVE RED
        if (typeof row[colNumber - 1] === 'number' && row[colNumber - 1] < 0) {
          cell.font = { color: { argb: 'FFFF0000' } };
        }
      });

      // ZEBRA ROWS (WORK MIX STYLE)
      if (excelRow.number % 2 === 0) {
        excelRow.eachCell((cell: any) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'F2F2F2' }
          };
        });
      }
    });

    //  COLUMN WIDTH
    worksheet.getColumn(1).width = 30;

    for (let i = 2; i <= 25; i++) {
      worksheet.getColumn(i).width = 14;
    }

    worksheet.addRow([]);

    // DOWNLOAD
    workbook.xlsx.writeBuffer().then((data: any) => {

      const blob = new Blob([data], {
        type: EXCEL_TYPE
      });

      FileSaver.saveAs(
        blob,
        'Unapplied Time Report_' + DATE_EXTENSION + EXCEL_EXTENSION
      );
    });
  }
  GetUnappliedReport(status: any, store: any, account: any, j: any) {

    // Reset
    if (status === '') {
      this.UnappliedData = [];
      this.NoData = '';
    }

    //  STEP 1: Past 3 Months (same as Angular 15 logic)
    this.pastThreeMonths = [];
    let currentDate = new Date(this.month);

    for (let i = 0; i < 3; i++) {
      const tempDate = new Date(currentDate);
      tempDate.setMonth(currentDate.getMonth() - i);
      this.pastThreeMonths.push(tempDate);
    }

    this.keys = this.pastThreeMonths.map(d =>
      this.shared.datePipe.transform(d, 'MMM yyyy')!
    );

    this.shared.spinner.show();

    const selectedDate = new Date(this.month);
    selectedDate.setDate(1);

    const baseObj = {
      STORE: this.storeIds.toString(),
      StartDate: this.shared.datePipe.transform(selectedDate, 'MM-dd-yyyy')
    };

    forkJoin([
      this.apiSrvc.postmethod(
        this.shared.common.routeEndpoint + 'GetUnAppliedTimeReport',
        { ...baseObj, Type: 'T' }
      ),
      this.apiSrvc.postmethod(
        this.shared.common.routeEndpoint + 'GetUnAppliedTimeReport',
        { ...baseObj, Type: 'D' }
      )
    ]).subscribe({

      next: ([totalRes, detailRes]: any) => {

        let combined: any[] = [];

        const parseData = (response: any[]) => {
          return response.map((item: any) => {

            let currentMonth: any = {};
            let past1Month: any = {};
            let past2Month: any = {};

            try { currentMonth = JSON.parse(item['@CurrenttMonth'] || '{}'); } catch { }
            try { past1Month = JSON.parse(item['@Past1Month'] || '{}'); } catch { }
            try { past2Month = JSON.parse(item['@Past2Month'] || '{}'); } catch { }

            return {
              Store_Name: item.Store_Name,

              [this.keys[0]]: {
                Mtd: currentMonth.Mtd ?? null,
                Labor_Sale_Per: currentMonth.Labor_Sale_Per ?? null,
                Budget: currentMonth.Budget ?? null,
                Diff: currentMonth.Diff ?? null
              },

              [this.keys[1]]: {
                Mtd: past1Month.Mtd ?? null,
                Labor_Sale_Per: past1Month.Labor_Sale_Per ?? null,
                Budget: past1Month.Budget ?? null,
                Diff: past1Month.Diff ?? null
              },

              [this.keys[2]]: {
                Mtd: past2Month.Mtd ?? null,
                Labor_Sale_Per: past2Month.Labor_Sale_Per ?? null,
                Budget: past2Month.Budget ?? null,
                Diff: past2Month.Diff ?? null
              }
            };
          });
        };

        if (totalRes?.status === 200 && totalRes.response?.length) {
          combined = [...combined, ...parseData(totalRes.response)];
        }

        if (detailRes?.status === 200 && detailRes.response?.length) {
          combined = [...combined, ...parseData(detailRes.response)];
        }

        const totalRow = combined.find(
          (x: any) => x.Store_Name?.toUpperCase() === 'REPORTS TOTAL'
        );

        const storeRows = combined.filter(
          (x: any) => x.Store_Name?.toUpperCase() !== 'REPORTS TOTAL'
        );

        this.UnappliedData = totalRow
          ? [...storeRows, totalRow]
          : storeRows;

        this.NoData = this.UnappliedData.length ? '' : 'No Data Found';

        this.shared.spinner.hide();
      },

      error: (err) => {
        console.error(err);
        this.toast.show('Error fetching data', 'danger', 'Error');
        this.shared.spinner.hide();
        this.NoData = 'No Data Found';
      }
    });
  }
  nestedsort(month: string, field: string) {

    this.isDesc = !this.isDesc;
    const direction = this.isDesc ? 1 : -1;

    const totalRow = this.UnappliedData.find(
      (x: any) => x.Store_Name?.toUpperCase() === 'REPORTS TOTAL'
    );

    let storeRows = this.UnappliedData.filter(
      (x: any) => x.Store_Name?.toUpperCase() !== 'REPORTS TOTAL'
    );

    storeRows.sort((a: any, b: any) => {

      const valA = a[month]?.[field] ?? 0;
      const valB = b[month]?.[field] ?? 0;

      if (valA < valB) return -1 * direction;
      if (valA > valB) return 1 * direction;
      return 0;
    });

    this.UnappliedData = totalRow
      ? [...storeRows, totalRow]
      : storeRows;
  }

  //apply filters
  viewreport() {
    this.activePopover = -1
    if (this.storeIds.length == 0) {
      this.toast.show('Please select atleast one store', 'warning', 'Warning');
    }

    else {
      this.expanddataarray = []

      if (this.storeIds != '') {
        this.GetUnappliedReport('', 0, '', '');
        this.setHeaderData()
      } else {
        this.NoData = '';
      }
    }
  }

  public inTheGreen(value: any): boolean {
    if (value >= 0) {
      return true;
    }
    return false;
  }

  togglePopover(popoverIndex: number) {
    this.activePopover = this.activePopover === popoverIndex ? -1 : popoverIndex;
  }
  ngAfterViewInit() {

    this.apiSrvc.getStores().subscribe((res: any) => {
      if (this.shared.common.pageName == 'Unapplied Time Report') {
        if (res.obj.storesData != undefined) {
          this.groupsArray = res.obj.storesData;
          this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
          this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_name : this.groupName = ''
          this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
          this.getStoresandGroupsValues()
        }
      }
    })
    this.excel = this.apiSrvc.getExportToExcelAllReports().subscribe((res) => {
      if (this.excel != undefined) {
        if (res.obj.title == 'Unapplied Time Report') {
          if (res.obj.state == true) {
            this.exportToExcel();
          }
        }
      }
    });
  }
}
