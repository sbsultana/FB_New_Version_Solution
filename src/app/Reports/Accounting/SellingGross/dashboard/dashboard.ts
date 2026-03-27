import { Component, ElementRef, HostListener, Injector, ViewChild } from '@angular/core';
import { Api } from '../../../../Core/Providers/Api/api';
import { common } from '../../../../common';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { FormsModule } from '@angular/forms';
import { NgbActiveModal, NgbModal } from '@ng-bootstrap/ng-bootstrap';
import { NgxSpinnerService } from 'ngx-spinner';
import { CurrencyPipe, DatePipe, formatDate } from '@angular/common';
import { Title } from '@angular/platform-browser';
import { Subscription } from 'rxjs';
import { environment } from '../../../../../environments/environment';
import * as ExcelJS from 'exceljs';
import FileSaver from 'file-saver';
import html2canvas from 'html2canvas';
import jsPDF from 'jspdf';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
const EXCEL_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
const EXCEL_EXTENSION = '.xlsx';
import numeral from 'numeral';
import { BsDatepickerConfig, BsDatepickerDirective, BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { Router } from '@angular/router';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { Stores } from '../../../../CommonFilters/stores/stores';

@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [SharedModule, BsDatepickerModule, Stores],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
})
export class Dashboard {
  subscription!: Subscription;
  subscriptionExcel!: Subscription;
  subscriptionReport!: Subscription;
  subscriptionPDF!: Subscription;
  subscriptionEmailPDF!: Subscription;
  FromDate: any = '';
  ToDate: any = '';
  SalesData: any = [];
  IndividualSalesGross: any = [];
  TotalSalesGross: any = [];
  TotalSortPosition: any = 'B';
  NoData: any = '';
  responcestatus: string = '';
  groups: any = [1];
  reportOpenSub!: Subscription;
  reportGetting!: Subscription;
  Pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;
  selectedGroups: any = [];
  storeName: any = ''
  month!: Date;
  DuplicatDate!: Date;
  previousMonth!: Date;

  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent) {
    const clickedInside = (event.target as HTMLElement).closest('.dropdown-toggle, .reportstores-card , .timeframe');
    if (!clickedInside) {
      this.activePopover = -1;
    }
  }

  selectedstorevalues: any = [];
  selectedstorename: any;
  selectedFilters: string[] = [];
  selectedLabel: string = '( All )';
  StoreName: any = 'All Stores';
  Filter: any = ['New', 'Used', 'Service', 'Parts', 'Detail'];
  stores: any = [];
  selectDate: Date = new Date();
  currentMonth: any = '';
  groupsArray: any = [];
  storename: any = ''
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  groupId: any = 0;
  storeIds: any = '0';
  storesFilterData: any = {
    'groupsArray': this.groupsArray, 'groupId': this.groupId, 'storesArray': this.stores, 'storeids': '1', 'type': 'M', 'others': 'N',
    'groupName': this.groupName, 'storename': this.storename, storecount: null, 'storedisplayname': this.storedisplayname
  };
  blocks: any = [
    // { label: 'All', value: 'All' },
    { label: 'Net to Sales', value: 'Net to Sales', tds: 3 },
    { label: 'Net to Gross', value: 'Net to Gross', tds: 3 },
    { label: 'Selling Gross', value: 'Selling Gross', tds: 8 },
  ]
  selectedBlocks: any = [];
  gridBlocks: any = [];
  departmentblocks: any = [];
  Dept: any = [
    // { label: 'All', value: 'All' },
    { label: 'New', value: 'New' },
    { label: 'Used', value: 'Used' },
    { label: 'Service', value: 'Service' },
    { label: 'Parts', value: 'Parts' },
  ]
  selectedDept: any = [
    { label: 'All', value: 'All' }
  ]
  date: any;
  constructor(
    private datepipe: DatePipe,
    public apiSrvc: Api,
    private ngbmodal: NgbModal,
    private ngbmodalActive: NgbActiveModal,
    private spinner: NgxSpinnerService,
    private title: Title,
    private comm: common,
    private router: Router,
    private toast: ToastService,
    private injector: Injector,
    public shared: Sharedservice,
  ) {
    const lastMonth = new Date();
    let today = new Date();
    if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
      // this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.ustores.split(',')
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
    if (today.getDate() < 5) {
      this.date = new Date(lastMonth.setMonth(lastMonth.getMonth() - 1));
    } else {
      this.date = new Date(lastMonth.setMonth(lastMonth.getMonth()));
    }
    this.title.setTitle(this.comm.titleName + '-Selling Gross');
    if (localStorage.getItem('Fav') != 'Y') {
      const data = {
        title: 'Selling Gross',
        path1: '',
        path2: '',
        path3: '',
        Month: this.date,
        stores: this.storeIds.toString(),
        store: this.storeIds,
        filter: this.Filter,
        groups: 1,
        count: 0,
      };
      this.apiSrvc.SetHeaderData({
        obj: data,
      });
    }
    this.selectedBlocks = [...this.blocks]
    this.gridBlocks = this.blocks.map(function (a: any) {
      return a.value;
    });
    this.selectedDept = [...this.Dept]
    console.log(this.selectedBlocks);
    this.departmentblocks = this.Dept.map(function (a: any) {
      return a.value;
    });
    this.GetData();
  }
  // Filters list
  filters: string[] = ['New', 'Used', 'Service', 'Parts', 'Detail'];
  activePopover: number | null = null;
  bsConfig: Partial<BsDatepickerConfig> = {
    dateInputFormat: 'MMMM/YYYY',
    minMode: 'month',
    maxDate: new Date()
  };
  monthPicker!: BsDatepickerDirective;
  openMonthPicker() {
    if (this.monthPicker) {
      this.monthPicker.show();
    }
  }

  togglePopover(index: number) {
    this.activePopover = this.activePopover === index ? null : index;
  }


  isSelected(filter: string): boolean {
    return this.selectedFilters.includes(filter);
  }

  toggleFilter(filter: string) {
    const index = this.selectedFilters.indexOf(filter);
    if (index >= 0) {
      this.selectedFilters.splice(index, 1);
    } else {
      this.selectedFilters.push(filter);
    }
    this.updateLabel();
  }
  getBackgroundColor(value: any, block: any) {
    if (block == 'New' || block == 'Used') {
      if (value == undefined) {
        return ''
      }
      else if (value < 45) {
        return '#e37474';
      }
      else {
        return '#6fb979';
      }
    } else if (block == 'Service') {
      if (value == undefined) {
        return ''
      }
      else if (value < 50) {
        return '#e37474';
      }
      else {
        return '#6fb979';
      }
    } else {
      if (value == undefined) {
        return ''
      }
      else if (value < 60) {
        return '#e37474';
      }
      else {
        return '#6fb979';
      }
    }

  }

  updateLabel() {
    if (!this.selectedFilters || this.selectedFilters.length === 0) {
      this.selectedLabel = '( Select )';
    } else if (this.selectedFilters.length === this.filters.length) {
      this.selectedLabel = '( All )';
    } else if (this.selectedFilters.length === 1) {
      this.selectedLabel = `( ${this.selectedFilters[0]} )`;
    } else {
      this.selectedLabel = `( Selected ${this.selectedFilters.length} )`;
    }
  }

  applyDateChange() {
    if (!this.storeIds || this.storeIds.length === 0) {
      this.toast.show('Please Select Atleast One Store', 'warning', 'Warning');
      return;
    }
    if (!this.selectedBlocks || this.selectedBlocks.length === 0) {
      this.toast.show('Please Select Atleast One Type', 'warning', 'Warning');
      return;
    }
    if (!this.selectedDept || this.selectedDept.length === 0) {
      this.toast.show('Please Select Atleast One Department', 'warning', 'Warning');
      return;
    }
    else {

      this.GetData();
      this.gridBlocks = this.selectedBlocks.map(function (a: any) {
        return a.value;
      });
      this.departmentblocks = this.selectedDept.map(function (a: any) {
        return a.value;
      });
      this.activePopover = null;
      this.isLoading = true;
    }
  }

  getColspan() {
    const ids2 = this.gridBlocks.map((obj: any) => obj);
    let commondata = this.blocks.filter((obj1: any) => ids2.includes(obj1.value));
    //  console.log(tds,'..................'); 
    let td = commondata.map(function (a: any) {
      return a.tds;
    });
    console.log(td, td.reduce((accumulator: any, currentValue: any) => accumulator + currentValue, 0), '.............');
    return td.reduce((accumulator: any, currentValue: any) => accumulator + currentValue, 0);
  }
  GetData() {
    this.IndividualSalesGross = [];
    this.spinner.show();
    let date = new Date()
    let endDate = new Date()
    let lastday: any = '01'
    if (this.selectDate.getMonth() == date.getMonth()) {
      endDate = new Date(date.setDate(date.getDate() - 1));
      lastday = ('0' + endDate.getDate()).slice(-2)
    }
    else {
      if (this.selectDate.getMonth() == 0) {
        lastday = '31'
      } else {
        var lastDayOfMonth = new Date(this.selectDate.getFullYear(), this.selectDate.getMonth() + 1, 0);
        lastday = ('0' + lastDayOfMonth.getDate()).slice(-2)
      }
    }
    this.DuplicatDate = this.selectDate
    let Prev_Date = new Date(this.selectDate)
    this.previousMonth = Prev_Date
    this.previousMonth.setMonth(this.previousMonth.getMonth());


    console.log(this.selectDate, '..............');

    const obj = {
      StartDate: this.datepipe.transform(this.selectDate, 'MM') + '-' + '01' + '-' + this.datepipe.transform(this.selectDate, 'YYYY'),
      EndDate: this.datepipe.transform(this.selectDate, 'MM') + '-' + lastday + '-' + this.datepipe.transform(this.selectDate, 'YYYY'),
      Stores: this.storeIds.toString(),
    };
    this.apiSrvc
      .postmethod(this.comm.routeEndpoint + 'GetSellingGrossV1', obj)
      .subscribe(
        (res) => {
          this.SalesData = [];
          if (res && res.response && res.response.length > 0) {
            this.spinner.hide();
            this.SalesData = res.response;

            console.log(res.response, '............');
            this.SalesData.some(function (x: any) {
              if (x.DetailData != undefined && x.DetailData != '' && x.DetailData != null) {
                x.DetailData = JSON.parse(x.DetailData);
                x.DetailData = x.DetailData.sort((a: any, b: any) => a.Slno - b.Slno)
              }
              if (x.YTD != undefined && x.YTD != '' && x.YTD != null) {
                x.YTD = JSON.parse(x.YTD);
              }
            });
            this.spinner.hide();
            this.NoData = '';
          } else {
            this.spinner.hide();
            this.NoData = 'No Data Found!!';
          }
        },
        (error) => {
          this.toast.show('502 Bad Gate Way Error', 'danger', 'Error');
          this.spinner.hide();
          this.NoData = 'No Data Found';
        }
      );
  }

  isLoading = true;
  formatMonth(date: Date): string {
    const year = date.getFullYear();
    const month = ('0' + (date.getMonth() + 1)).slice(-2);
    return `${year}-${month}`;
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
    if (direction == -1) {
      this.TotalSortPosition = 'T';
    } else {
      this.TotalSortPosition = 'B';
    }
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
  SFstate: any;
  ngAfterViewInit(): void {
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.shared.common.pageName == 'Selling Gross') {
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
        if (res.obj.title == 'Selling Gross') {
          if (res.obj.state == true) {
            this.exportToExcel();
          }
        }
      }
    });
    this.email = this.apiSrvc.getExportToEmailPDFAllReports().subscribe((res) => {
      if (this.email != undefined) {
        if (res.obj.title == 'Selling Gross') {
          if (res.obj.stateEmailPdf == true) {
            this.sendEmailData(res.obj.Email, res.obj.notes, res.obj.from);
          }
        }
      }
    });
    this.print = this.apiSrvc.getExportToPrintAllReports().subscribe((res) => {
      if (this.print != undefined) {
        if (res.obj.title == 'Selling Gross') {
          if (res.obj.statePrint == true) {
            this.GetPrintData();
          }
        }
      }
    });

    this.Pdf = this.apiSrvc.getExportToPDFAllReports().subscribe((res) => {
      if (this.Pdf != undefined) {
        console.log('PDF');
        if (res.obj.title == 'Selling Gross') {
          console.log('PDF Selling Gross');
          if (res.obj.statePDF == true) {
            this.generatePDF();
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

    // this.setHeaderData();
    // this.GetData();

  }
  StoresData(data: any) {
    this.storeIds = data.storeids;
    this.groupId = data.groupId;
    this.storename = data.storename;
    this.groupName = data.groupName;
    this.storecount = data.storecount;
    this.storedisplayname = data.storedisplayname;
  }
  multipleorsingle(block: any, e: any) {
    if (block == 'B') {
      if (e == 'All') {
        this.selectedBlocks.length == this.blocks.length ? this.selectedBlocks = [] : this.selectedBlocks = [...this.blocks]
      } else {
        const index = this.selectedBlocks.findIndex((i: any) => i == e);
        if (index >= 0) {
          this.selectedBlocks.splice(index, 1);
        } else {
          this.selectedBlocks.push(e);
        }
      }
    } if (block == 'D') {
      if (e == 'All') {
        this.selectedDept.length == this.Dept.length ? this.selectedDept = [] : this.selectedDept = [...this.Dept]
      } else {
        const index = this.selectedDept.findIndex((i: any) => i == e);
        if (index >= 0) {
          this.selectedDept.splice(index, 1);
        } else {
          this.selectedDept.push(e);
        }
      }
    }


  }

  getSelectedStoreNames(): string {
    if (!this.storeIds || this.storeIds.length === 0) return '';

    const ids = this.storeIds.toString().split(',');

    const selectedStores = this.stores.filter((s: any) =>
      ids.includes(s.ID.toString())
    );

    return selectedStores.map((s: any) => s.storename).join(', ');
  }
  getReportFilters(): { title: string; filters: any[] } {
    return {
      title: 'Selling Gross',
      filters: [
        {
          label: 'Store',
          value: this.getSelectedStoreNames() || 'All Stores'
        },
        {
          label: 'Group',
          value: this.groupName || ''
        },
        {
          label: 'Month',
          value: this.datepipe.transform(this.selectDate, 'MMMM yyyy')
        },
        {
          label: 'Type',
          value: this.selectedBlocks.length === this.blocks.length
            ? 'All'
            : this.selectedBlocks.map((b: any) => b.value).join(', ')
        },
        {
          label: 'Department',
          value: this.selectedDept.length === this.Dept.length
            ? 'All'
            : this.selectedDept.map((d: any) => d.value).join(', ')
        }
      ]
    };
  }
  addExcelFiltersSection(worksheet: any): number {
    let rowCount = 0;

    const report = this.getReportFilters();

    /*  TITLE (LEFT ALIGNED) */
    const titleRow = worksheet.addRow([report.title]);
    titleRow.font = { bold: true, size: 14 };
    worksheet.mergeCells(`A${rowCount + 1}:G${rowCount + 1}`);
    titleRow.alignment = { horizontal: 'left', vertical: 'middle' };
    rowCount++;

    /* FILTERS */
    report.filters.forEach((filter: any) => {
      const row = worksheet.addRow([`${filter.label}:`, filter.value]);
      row.getCell(1).font = { bold: true };
      rowCount++;
    });

    /* SPACE */
    worksheet.addRow([]);
    rowCount++;

    return rowCount;
  }
  exportToExcel(): void {

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Selling Gross');

    /* ================= FILTER SECTION ================= */
    const filterRowCount = this.addExcelFiltersSection(worksheet);

    /* ================= FORMAT FUNCTION ================= */
    const formatRow = (row: any) => {
      row.eachCell((cell: any, colNumber: number) => {

        cell.border = {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' }
        };

        if (colNumber === 1) {
          cell.alignment = { horizontal: 'left', vertical: 'middle' };
        } else {
          cell.alignment = { horizontal: 'right', vertical: 'middle' };
        }

        if (!cell.value || cell.value === '-') {
          cell.value = '-';
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
        }
      });

      /* ===== ZEBRA ===== */
      if (row.number % 2 === 0) {
        row.eachCell((cell: any) => {
          if (!cell.fill) {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'F5F7FA' }
            };
          }
        });
      }
    };

    /* ================= GROUP HEADER ================= */
    const selectedMonth = this.datepipe.transform(this.selectDate, 'MMM yyyy');

    const groupHeader: any[] = [selectedMonth || ''];
    if (this.gridBlocks.includes('Net to Sales')) {
      groupHeader.push('Net to Sales %', '');
    }
    if (this.gridBlocks.includes('Net to Gross')) {
      groupHeader.push('Net to Gross %', '');
    }
    if (this.gridBlocks.includes('Selling Gross')) {
      groupHeader.push('Selling Gross', '', '', '', 'YTD', '');
    }

    const groupHeaderRow = worksheet.addRow(groupHeader);

    let colIndex = 2;

    if (this.gridBlocks.includes('Net to Sales')) {
      worksheet.mergeCells(groupHeaderRow.number, colIndex, groupHeaderRow.number, colIndex + 1);
      colIndex += 2;
    }
    if (this.gridBlocks.includes('Net to Gross')) {
      worksheet.mergeCells(groupHeaderRow.number, colIndex, groupHeaderRow.number, colIndex + 1);
      colIndex += 2;
    }
    if (this.gridBlocks.includes('Selling Gross')) {
      worksheet.mergeCells(groupHeaderRow.number, colIndex, groupHeaderRow.number, colIndex + 3);
      colIndex += 4;
      worksheet.mergeCells(groupHeaderRow.number, colIndex, groupHeaderRow.number, colIndex + 1);
    }

    groupHeaderRow.eachCell((cell: any) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF0554EF' } // ✅ DARK BLUE
      };
      cell.font = { bold: true, color: { argb: 'FFFFFF' }, name: 'Calibri' };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
    });

    /* ================= MAIN HEADER ================= */
    const headerRow = [''];
    const month = this.datepipe.transform(this.previousMonth, 'MMM yy') || '';
    if (this.gridBlocks.includes('Net to Sales')) {
      headerRow.push(
        month,
        'Jan-' + month
      );
    }
    if (this.gridBlocks.includes('Net to Gross')) {
      headerRow.push(
        month,
        'Jan-' + month
      );
    }
    if (this.gridBlocks.includes('Selling Gross')) {
      headerRow.push(
        month + ' Gross',
        month,
        'Jan-' + month + ' Gross',
        'Jan-' + month,
        'YTD' + ' Gross',
        'YTD'
      );
    }

    const excelHeader = worksheet.addRow(headerRow);

    excelHeader.eachCell((cell: any) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF4584FF' } // ✅ LIGHT BLUE
      };
      cell.font = { bold: true, color: { argb: 'FFFFFF' }, name: 'Calibri' };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };

      cell.border = {
        top: { style: 'thin' },
        bottom: { style: 'thin' },
        left: { style: 'thin' },
        right: { style: 'thin' }
      };
    });

    /* ================= DATA ================= */

    for (const info of this.SalesData) {

      const row = [info.AS_DEALERNAME ?? '--'];

      if (this.gridBlocks.includes('Net to Sales')) {
        row.push(info.NS_MTD || '-', info.NS_YTD || '-');
      }

      if (this.gridBlocks.includes('Net to Gross')) {
        row.push(info.NG_MTD || '-', info.NG_YTD || '-');
      }

      if (this.gridBlocks.includes('Selling Gross')) {
        row.push('', '', '', '');
      }

      const dealerRow = worksheet.addRow(row);

      // ✅ Apply full row background color
      dealerRow.eachCell((cell: any) => {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFD9E7FF' } // D9E7FF
        };
      });

      // apply borders + alignment (after color)
      formatRow(dealerRow);

      for (const dept of info.DetailData) {

        if (this.departmentblocks.includes(dept.TitleValue)) {

          const nestedRow = ['   ' + dept.TitleValue];

          if (this.gridBlocks.includes('Net to Sales')) nestedRow.push('', '');
          if (this.gridBlocks.includes('Net to Gross')) nestedRow.push('', '');

          if (this.gridBlocks.includes('Selling Gross')) {
            nestedRow.push(
              dept.MTD_Gross || '-',
              dept.MTD || '-',
              dept.YTD_Gross || '-',
              dept.YTD || '-',
              dept.YTD_Gross_YTD || '-',
              dept.YTD_YTD || '-'
            );
          }

          const rowRef = worksheet.addRow(nestedRow);
          formatRow(rowRef);

          const add = (cell: any, type: 'pct' | 'cur') => {
            if (!cell || cell.value === null || cell.value === undefined || cell.value === '-') return;

            let val = cell.value.toString().replace('%', '').replace('$', '');

            if (val === '') return;

            if (type === 'cur') {
              cell.value = '$' + val;
            } else {
              const num = parseFloat(val);

              if (!isNaN(num)) {
                cell.value = num.toFixed(1) + '%'; // ✅ ONLY DISPLAY FIX
              }
            }
          };
          let col = 2; // start after dealer column

          // Net to Sales → %
          if (this.gridBlocks.includes('Net to Sales')) {
            add(rowRef.getCell(col), 'pct');
            add(rowRef.getCell(col + 1), 'pct');
            col += 2;
          }

          // Net to Gross → %
          if (this.gridBlocks.includes('Net to Gross')) {
            add(rowRef.getCell(col), 'pct');
            add(rowRef.getCell(col + 1), 'pct');
            col += 2;
          }

          // Selling Gross
          if (this.gridBlocks.includes('Selling Gross')) {

            // SAFE: check value exists before applying
            add(rowRef.getCell(col), 'cur');       // Mar Gross $
            add(rowRef.getCell(col + 1), 'pct');   // Mar %
            add(rowRef.getCell(col + 2), 'cur');   // Jan-Mar Gross $
            add(rowRef.getCell(col + 3), 'pct');   // Jan-Mar %
            add(rowRef.getCell(col + 4), 'cur');   // YTD Gross $
            add(rowRef.getCell(col + 5), 'pct');   // YTD %
          }
          /* ===== KEEP YOUR COLOR LOGIC ===== */

          // ✅ DEFINE FIRST
          const applyColor = (cell: any, block: any) => {
            if (!cell || cell.value === null || cell.value === undefined || cell.value === '-') return;

            let val = typeof cell.value === 'number'
              ? cell.value
              : parseFloat(cell.value.toString().replace('%', ''));

            if (isNaN(val)) return;

            const bg = this.getBackgroundColor(val, block);

            if (!bg) return;

            const hex = bg.replace('#', '').toUpperCase();
            const argb = 'FF' + hex;

            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb }
            };
          };

          // ✅ THEN USE
          if (this.gridBlocks.includes('Selling Gross')) {

            let i = 2;

            if (this.gridBlocks.includes('Net to Sales')) i += 2;
            if (this.gridBlocks.includes('Net to Gross')) i += 2;

            applyColor(rowRef.getCell(i + 1), dept.TitleValue); // MTD %
            applyColor(rowRef.getCell(i + 3), dept.TitleValue); // Jan-Mar %
            applyColor(rowRef.getCell(i + 5), dept.TitleValue); // YTD %
          }
        }
      }
    }

    /* ================= AUTO WIDTH ================= */
    worksheet.columns.forEach((col: any, index: number) => {

      // Dealer column
      if (index === 0) {
        col.width = 30;
      }

      // % columns
      else {
        col.width = 20;
      }

    });
    /* ================= EXPORT ================= */
    workbook.xlsx.writeBuffer().then((buffer: any) => {
      const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      FileSaver.saveAs(blob, `Selling_Gross`);
    });

  }



  GetPrintData() {
    window.print();
  }

  generatePDF() {
    this.spinner.show();
    const printContents = document.getElementById('SellingGross')!.innerHTML;
    const iframe = document.createElement('iframe');

    iframe.style.position = 'absolute';
    iframe.style.width = '0px';
    iframe.style.height = '0px';
    iframe.style.border = 'none';

    document.body.appendChild(iframe);
    const doc = iframe.contentWindow?.document;
    if (!doc) {
      console.error('Failed to create iframe document');
      return;
    }

    doc.open();
    doc.write(`
    <html>
            <head>
              <title>Selling Gross</title>
                 <style>
                 @font-face {
                  font-family: 'GothamBookRegular';
                  src: url('assets/fonts/Gotham\ Book\ Regular.otf') format('otf'), /* Chrome 6+, Firefox 3.6+, IE 9+, Safari 5.1+ */
                       url('assets/fonts/Gotham\ Book\ Regular.otf') format('opentype'); /* Chrome 4+, Firefox 3.5, Opera 10+, Safari 3—5 */
                }
                @font-face {
                  font-family: 'Roboto';
                  src: url('assets/fonts/Roboto-Regular.ttf') format('ttf'), /* Chrome 6+, Firefox 3.6+, IE 9+, Safari 5.1+ */
                       url('assets/fonts/Roboto-Regular.ttf') format('truetype'); /* Chrome 4+, Firefox 3.5, Opera 10+, Safari 3—5 */
                }
                @font-face {
                  font-family: 'RobotoBold';
                  src: url('assets/fonts/Roboto-Bold.ttf') format('ttf'), /* Chrome 6+, Firefox 3.6+, IE 9+, Safari 5.1+ */
                       url('assets/fonts/Roboto-Bold.ttf') format('truetype'); /* Chrome 4+, Firefox 3.5, Opera 10+, Safari 3—5 */
                }
  
                
                
                .negative {
                  color: red;
                }
                .HideItems{
                  display:none;
                 }            
        
  .backgrossbtn {
                  float: left;
                  font-size: 11px;
                  background-color: #2d91f1;
                  color: white;
                  padding: 4px;
                  border: 1px solid #2d91f1;
                  cursor: pointer;
                  border-radius: 3px;
              }
                .Selectedrow {
                  background-color: #d3ecff !important;
                  color: #000 !important;
              }
  
                .justify-content-between {
                  justify-content: space-between !important;
              }
              .d-flex {
                  display: flex !important;
              }
                .bg-white {
                  background: #ffffff !important;
              }
                .negative {
                  color: red;
                }
                
                .negative {
                  color: red;
                }
                .performance-scorecard .table>:not(:first-child) {
                  border-top: 0px solid #ffa51a
                }
                .performance-scorecard .table {
                  text-align: center;
                  text-transform: capitalize;
                  border: transparent;
            
                  width: 100%;
                }
                .performance-scorecard .table th,
                .performance-scorecard .table td{
                  white-space: nowrap;
                  vertical-align: top;
                }
                .performance-scorecard .table th:first-child,
                .performance-scorecard .table td:first-child {
                  position: sticky;
                  left: 0;
                  z-index: 1;
                  background-color: #337ab7;
                }
                .performance-scorecard .table tr:nth-child(odd) td:first-child,
                .performance-scorecard .table tr:nth-child(odd) td:nth-child(2) {
                  // background-color: #e9ecef;
                  background-color: #ffffff;
                }
                .performance-scorecard .table tr:nth-child(even) td:first-child,
                .performance-scorecard .table tr:nth-child(even) td:nth-child(2) {
                  background-color: #ffffff;
                }
                .performance-scorecard .table tr:nth-child(odd) {
                  // background-color: #e9ecef;
                  background-color: #ffffff;
        
                }
                .performance-scorecard .table tr:nth-child(even) {
                  background-color: #ffffff;
                }
                .performance-scorecard .table .spacer {
                  // width: 50px !important;
                  background-color: #cfd6de !important;
                  border-left: 1px solid #cfd6de !important;
                  border-bottom: 1px solid #cfd6de !important;
                  border-top: 1px solid #cfd6de !important;
                 }
                .performance-scorecard .table .hidden {
                  display: none !important;
                }
                .performance-scorecard .table .bdr-rt {
                  border-right: 1px solid #abd0ec;
                }
                .performance-scorecard .table thead {
                  position: sticky;
                  top: 0;
                  z-index: 99;
                  font-family: 'FaktPro-Bold';
                  font-size: 0.8rem;
                }
                .performance-scorecard .table thead th {
                  padding: 5px 10px;
                  margin: 0px;
                }
                .performance-scorecard .table thead .bdr-btm {
                  border-bottom: #005fa3;
                }
                .performance-scorecard .table thead tr:nth-child(1) {
                  background-color: #fff !important;
                  color: #000;
                  text-transform: uppercase;
                  border-bottom: #cfd6de;
                }
                .performance-scorecard .table thead tr:nth-child(2) {
                  background-color: #fff !important;
                  color: #000;
                  text-transform: uppercase;
                  border-bottom: #cfd6de;
                  box-shadow: inset 0 1px 0 0 #cfd6de;
                }
                .performance-scorecard .table thead tr:nth-child(2) th :nth-child(1) {
                  background-color: #337ab7 !important;
                  color: #fff;
                }
                .performance-scorecard .table thead tr:nth-child(3) {
      
                  background-color: #fff !important;
                  color: #000;
                  text-transform: uppercase;
                  border-bottom: #cfd6de;
                  box-shadow: inset 0 1px 0 0 #cfd6de;
                }
                .performance-scorecard .table thead tr:nth-child(3) th :nth-child(1) {
                  background-color: #337ab7 !important;
                  color: #fff;
                }
                .performance-scorecard .table tbody {
                  font-family: 'FaktPro-Normal';
                  font-size: .9rem;
                }
                .performance-scorecard .table tbody td {
                  padding: 2px 10px;
                  margin: 0px;
                  border: 1px solid #cfd6de
                }
      
                .performance-scorecard .table tbody tr {
                  border-bottom: 1px solid #37a6f8;
                  border-left: 1px solid #37a6f8
                }
      
                .performance-scorecard .table tbody td:first-child {
                  text-align: start;
                  box-shadow: inset -1px 0 0 0 #cfd6de;
                }
                .performance-scorecard .table tbody tr td:not(:first-child) {
                  text-align: right !important;
      
                }
                .performance-scorecard .table tbody .sub-title {
                  font-size: .8rem !important;
                }
                .performance-scorecard .table tbody .sub-subtitle{
                  font-size: .7rem !important;
                }
                .performance-scorecard .table tbody .text-bold {
                  font-family: 'FaktPro-Bold';
                }
                .performance-scorecard .table tbody .darkred-bg {
                  background-color: #282828 !important;
                  color: #fff;
                }
                .performance-scorecard .table tbody .lightblue-bg {
                  background-color: #646e7a !important;
                  color: #fff;
                }
                .performance-scorecard .table tbody .gold-bg {
                  background-color: #ffa51a;
                  color: #fff;
                }
                 </style>
            </head>
        <body id='content'>
        ${printContents}
        </body>
          </html>`);
    doc.close();

    const div = doc.getElementById('content');
    const options = {
      logging: true,
      allowTaint: false,
      useCORS: true,
    };
    if (!div) {
      console.error('Element not found');
      return;
    }
    html2canvas(div, options)
      .then((canvas) => {
        let imgWidth = 285;
        let pageHeight = 206;
        let imgHeight = (canvas.height * imgWidth) / canvas.width;
        let heightLeft = imgHeight;
        const contentDataURL = canvas.toDataURL('image/png');
        let pdfData = new jsPDF('l', 'mm', 'a4', true);
        let position = 5;

        function addExtraDataToPage(pdf: any, extraData: any, positionY: any) {
          pdf.text(extraData, 10, positionY);
        }

        function addPageAndImage(pdf: any, contentDataURL: any, position: any) {
          pdf.addPage();
          pdf.addImage(
            contentDataURL,
            'PNG',
            5,
            position,
            imgWidth,
            imgHeight,
            undefined,
            'FAST'
          );
        }

        pdfData.addImage(
          contentDataURL,
          'PNG',
          5,
          position,
          imgWidth,
          imgHeight,
          undefined,
          'FAST'
        );
        addExtraDataToPage(pdfData, '', position + imgHeight + 10);
        heightLeft -= pageHeight;

        while (heightLeft >= 0) {
          position = heightLeft - imgHeight;
          addPageAndImage(pdfData, contentDataURL, position);
          addExtraDataToPage(pdfData, '', position + imgHeight + 10);
          heightLeft -= pageHeight;
        }

        return pdfData;
      })
      .then((doc) => {
        doc.save('sellingreport.pdf');
        // popupWin.close();
        this.spinner.hide();
      });
  }

  sendEmailData(Email: any, notes: any, from: any) {
    this.spinner.show();
    const printContents = document.getElementById('SellingGross')!.innerHTML;
    const iframe = document.createElement('iframe');

    // Make the iframe invisible
    iframe.style.position = 'absolute';
    iframe.style.width = '0px';
    iframe.style.height = '0px';
    iframe.style.border = 'none';

    document.body.appendChild(iframe);

    const doc = iframe.contentWindow?.document;
    if (!doc) {
      console.error('Failed to create iframe document');
      return;
    }

    doc.open();
    doc.write(`
          <html>
              <head>
                  <title>Selling Gross</title>
                  <style>
                  @font-face {
                    font-family: 'GothamBookRegular';
                    src: url('assets/fonts/Gotham\ Book\ Regular.otf') format('otf'), /* Chrome 6+, Firefox 3.6+, IE 9+, Safari 5.1+ */
                         url('assets/fonts/Gotham\ Book\ Regular.otf') format('opentype'); /* Chrome 4+, Firefox 3.5, Opera 10+, Safari 3—5 */
                  }
                  @font-face {
                    font-family: 'Roboto';
                    src: url('assets/fonts/Roboto-Regular.ttf') format('ttf'), /* Chrome 6+, Firefox 3.6+, IE 9+, Safari 5.1+ */
                         url('assets/fonts/Roboto-Regular.ttf') format('truetype'); /* Chrome 4+, Firefox 3.5, Opera 10+, Safari 3—5 */
                  }
                  @font-face {
                    font-family: 'RobotoBold';
                    src: url('assets/fonts/Roboto-Bold.ttf') format('ttf'), /* Chrome 6+, Firefox 3.6+, IE 9+, Safari 5.1+ */
                         url('assets/fonts/Roboto-Bold.ttf') format('truetype'); /* Chrome 4+, Firefox 3.5, Opera 10+, Safari 3—5 */
                  }
                  .negative {
                    color: red;
                  }
                  .HideItems{
                    display:none;
                   }            
           
  .backgrossbtn {
                    float: left;
                    font-size: 11px;
                    background-color: #2d91f1;
                    color: white;
                    padding: 4px;
                    border: 1px solid #2d91f1;
                    cursor: pointer;
                    border-radius: 3px;
                }
                  .Selectedrow {
                    background-color: #d3ecff !important;
                    color: #000 !important;
                }
  
                  .justify-content-between {
                    justify-content: space-between !important;
                }
                .d-flex {
                    display: flex !important;
                }
                  .bg-white {
                    background: #ffffff !important;
                }
                  .negative {
                    color: red;
                  }
                  
                  .negative {
                    color: red;
                  }
                  .performance-scorecard .table>:not(:first-child) {
                    border-top: 0px solid #ffa51a
                  }
                  .performance-scorecard .table {
                    text-align: center;
                    text-transform: capitalize;
                    border: transparent;
              
                    width: 100%;
                  }
                  .performance-scorecard .table th,
                  .performance-scorecard .table td{
                    white-space: nowrap;
                    vertical-align: top;
                  }
                  .performance-scorecard .table th:first-child,
                  .performance-scorecard .table td:first-child {
                    position: sticky;
                    left: 0;
                    z-index: 1;
                    background-color: #337ab7;
                  }
                  .performance-scorecard .table tr:nth-child(odd) td:first-child,
                  .performance-scorecard .table tr:nth-child(odd) td:nth-child(2) {
                    // background-color: #e9ecef;
                    background-color: #ffffff;
                  }
                  .performance-scorecard .table tr:nth-child(even) td:first-child,
                  .performance-scorecard .table tr:nth-child(even) td:nth-child(2) {
                    background-color: #ffffff;
                  }
                  .performance-scorecard .table tr:nth-child(odd) {
                    // background-color: #e9ecef;
                    background-color: #ffffff;
          
                  }
                  .performance-scorecard .table tr:nth-child(even) {
                    background-color: #ffffff;
                  }
                  .performance-scorecard .table .spacer {
                    // width: 50px !important;
                    background-color: #cfd6de !important;
                    border-left: 1px solid #cfd6de !important;
                    border-bottom: 1px solid #cfd6de !important;
                    border-top: 1px solid #cfd6de !important;
                   }
                  .performance-scorecard .table .hidden {
                    display: none !important;
                  }
                  .performance-scorecard .table .bdr-rt {
                    border-right: 1px solid #abd0ec;
                  }
                  .performance-scorecard .table thead {
                    position: sticky;
                    top: 0;
                    z-index: 99;
                    font-family: 'FaktPro-Bold';
                    font-size: 0.8rem;
                  }
                  .performance-scorecard .table thead th {
                    padding: 5px 10px;
                    margin: 0px;
                  }
                  .performance-scorecard .table thead .bdr-btm {
                    border-bottom: #005fa3;
                  }
                  .performance-scorecard .table thead tr:nth-child(1) {
                    background-color: #fff !important;
                    color: #000;
                    text-transform: uppercase;
                    border-bottom: #cfd6de;
                  }
                  .performance-scorecard .table thead tr:nth-child(2) {
                    background-color: #fff !important;
                    color: #000;
                    text-transform: uppercase;
                    border-bottom: #cfd6de;
                    box-shadow: inset 0 1px 0 0 #cfd6de;
                  }
                  .performance-scorecard .table thead tr:nth-child(2) th :nth-child(1) {
                    background-color: #337ab7 !important;
                    color: #fff;
                  }
                  .performance-scorecard .table thead tr:nth-child(3) {
        
                    background-color: #fff !important;
                    color: #000;
                    text-transform: uppercase;
                    border-bottom: #cfd6de;
                    box-shadow: inset 0 1px 0 0 #cfd6de;
                  }
                  .performance-scorecard .table thead tr:nth-child(3) th :nth-child(1) {
                    background-color: #337ab7 !important;
                    color: #fff;
                  }
                  .performance-scorecard .table tbody {
                    font-family: 'FaktPro-Normal';
                    font-size: .9rem;
                  }
                  .performance-scorecard .table tbody td {
                    padding: 2px 10px;
                    margin: 0px;
                    border: 1px solid #cfd6de
                  }
        
                  .performance-scorecard .table tbody tr {
                    border-bottom: 1px solid #37a6f8;
                    border-left: 1px solid #37a6f8
                  }
        
                  .performance-scorecard .table tbody td:first-child {
                    text-align: start;
                    box-shadow: inset -1px 0 0 0 #cfd6de;
                  }
                  .performance-scorecard .table tbody tr td:not(:first-child) {
                    text-align: right !important;
        
                  }
                  .performance-scorecard .table tbody .sub-title {
                    font-size: .8rem !important;
                  }
                  .performance-scorecard .table tbody .sub-subtitle{
                    font-size: .7rem !important;
                  }
                  .performance-scorecard .table tbody .text-bold {
                    font-family: 'FaktPro-Bold';
                  }
                  .performance-scorecard .table tbody .darkred-bg {
                    background-color: #282828 !important;
                    color: #fff;
                  }
                  .performance-scorecard .table tbody .lightblue-bg {
                    background-color: #646e7a !important;
                    color: #fff;
                  }
                  .performance-scorecard .table tbody .gold-bg {
                    background-color: #ffa51a;
                    color: #fff;
                  }
  
                  </style>
              </head>
              <body id='content'>
                  ${printContents}
              </body>
          </html>
      `);

    doc.close();

    const div = doc.getElementById('content');
    if (!div) {
      console.error('Element not found');
      return;
    }
    const options = {
      logging: true,
      allowTaint: false,
      useCORS: true,

    };

    html2canvas(div, options)
      .then((canvas) => {
        let imgWidth = 285;
        let pageHeight = 206;
        let imgHeight = (canvas.height * imgWidth) / canvas.width;
        let heightLeft = imgHeight;
        const contentDataURL = canvas.toDataURL('image/png');
        let pdfData = new jsPDF('l', 'mm', 'a4', true);
        let position = 5;
        pdfData.addImage(
          contentDataURL,
          'PNG',
          5,
          position,
          imgWidth,
          imgHeight,
          undefined,
          'FAST'
        );
        heightLeft -= pageHeight;
        while (heightLeft >= 0) {
          position = heightLeft - imgHeight;
          pdfData.addPage();
          pdfData.addImage(
            contentDataURL,
            'PNG',
            5,
            position,
            imgWidth,
            imgHeight,
            undefined,
            'FAST'
          );
          heightLeft -= pageHeight;
        }

        const pdfBlob = pdfData.output('blob');
        const pdfFile = this.blobToFile(pdfBlob, 'SalesGross.pdf');
        const formData = new FormData();
        formData.append('to_email', Email);
        formData.append('subject', 'Selling Gross');
        formData.append('file', pdfFile);
        formData.append('notes', notes);
        formData.append('from', from);
        this.apiSrvc.postmethod(this.comm.routeEndpoint + 'mail', formData).subscribe(
          (res: any) => {
            console.log('Response:', res);
            if (res.status === 200) {
              // this.toast.show(res.response);
              this.toast.success(res.response);
            } else {
              this.toast.show('Invalid Details','danger','Error');
            }
          },
          (error) => {
            console.error('Error:', error);
          }
        );
      })
      .catch((error) => {
        console.error('html2canvas error:', error);
      })
      .finally(() => {
        this.spinner.hide();
        // popupWin.close();
      });
  }
  public blobToFile = (theBlob: Blob, fileName: string): File => {
    return new File([theBlob], fileName, {
      lastModified: new Date().getTime(),
      type: theBlob.type,
    });
  };
}
