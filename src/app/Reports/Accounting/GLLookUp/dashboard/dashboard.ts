import { Component, Injector, HostListener } from '@angular/core';
import { Api } from '../../../../Core/Providers/Api/api';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { NgbActiveModal, NgbModal } from '@ng-bootstrap/ng-bootstrap';
import { Subscription } from 'rxjs';
import { Workbook } from 'exceljs';
import FileSaver from 'file-saver';
import html2canvas from 'html2canvas';
import jsPDF from 'jspdf';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
const EXCEL_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
const EXCEL_EXTENSION = '.xlsx';
import numeral from 'numeral';

import { BsDatepickerConfig, BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { Stores } from '../../../../CommonFilters/stores/stores';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';

@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [SharedModule, BsDatepickerModule, Stores],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
})
export class Dashboard {

  GLData: any = [];
  grandTotal: number = 0;
  accountNumber: string = '';
  NoData = '';
  actionType: any = 'N'
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
  activePopover: number = -1;

  month!: Date;
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
  constructor(public apiSrvc: Api, private ngbmodalActive: NgbActiveModal,
    private toast: ToastService, private injector: Injector, public shared: Sharedservice,) {

    let today = new Date();
    let enddate = new Date(today.setDate(today.getDate() - 1));
    this.month = new Date(enddate.setMonth(enddate.getMonth() - 1))
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
    this.shared.setTitle(this.shared.common.titleName + '-GL Lookup')

    this.setHeaderData()
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
  setHeaderData() {
    const data = {
      title: 'GL Lookup',
      Month: this.month,
      store: this.storeIds,
      groups: this.groupId,
    };
    this.apiSrvc.SetHeaderData({
      obj: data,
    });
  }
  searchGLData() {
    this.GLData = [];
    const acc = (this.accountNumber || '').trim();


    this.shared.spinner.show();
    const obj = {
      Store: this.storeIds.toString(),
      Account_Number: acc,
      Date: this.shared.datePipe.transform(this.month, 'yyyy-MM-dd'),

    };

    this.shared.api.postmethod(this.shared.common.routeEndpoint + 'GetGLLookup', obj).subscribe(
      (res: any) => {
        this.shared.spinner.hide();
        if (res.status === 200 && res.response.length > 0) {
          this.GLData = res.response;
          this.NoData = '';
          this.grandTotal = this.GLData.reduce((sum: any, r: any) => {
            return sum + (r.postingamount ? Number(r.postingamount) : 0);
          }, 0);

        } else {
          this.GLData = [];
          this.NoData = 'No Data Found!!';
        }
      },
      () => {
        this.shared.spinner.hide();
        this.toast.show('Error fetching GL data', 'danger', 'Error');
        this.NoData = 'No Data Found!!';

      }
    );
  }

  expandedGLRows: any[] = [];

  expandGL(i: number, row: any) {
    if (this.expandedGLRows.includes(i)) {
      this.expandedGLRows = this.expandedGLRows.filter(x => x !== i);
      return;
    }
    this.expandedGLRows.push(i);
    if (!row.childRows) {
      try {
        row.childRows = JSON.parse(row.AccountDetails);
      } catch {
        row.childRows = [];
      }
    }
  }


  isDesc: boolean = false;
  column: string = '';

  sort(property: any, data: any, block: any) {

    this.isDesc = !this.isDesc;  // toggle ASC/DESC
    this.column = property;
    let direction = this.isDesc ? 1 : -1;

    // ===================== GL MAIN BLOCK ("M") ===================== //
    if (block === 'M') {
      let fakedata = [...data];
      fakedata.sort((a: any, b: any) => {
        if (a[property] < b[property]) return -1 * direction;
        if (a[property] > b[property]) return 1 * direction;
        return 0;
      });
      this.GLData = [...fakedata];
      this.expandedGLRows = []
      return;
    }

    // ===================== GL CHILD TABLE BLOCK ("GL") ===================== //
    if (block === 'GL') {
      if (!data || data.length === 0) return;
      data.sort((a: any, b: any) => {
        let valA = a[property];
        let valB = b[property];
        if (property === 'accountingdate') {
          valA = new Date(valA);
          valB = new Date(valB);
        }

        valA = valA ?? '';
        valB = valB ?? '';

        if (valA < valB) return -1 * direction;
        if (valA > valB) return 1 * direction;
        return 0;
      });

      return;
    }

    // ===================== DEFAULT SIMPLE SORT ===================== //
    data.sort((a: any, b: any) => {
      if (a[property] < b[property]) return -1 * direction;
      if (a[property] > b[property]) return 1 * direction;
      return 0;
    });
  }
  viewreport() {
    console.log(this.storeIds);
    this.activePopover = -1
    if (this.storeIds.length == 0) {
      this.toast.show('Please select atleast one store', 'warning', 'Warning');
    }
    if (!this.accountNumber) {
      this.toast.show('Please enter an account number', 'warning', 'Warning');
      return;
    }

    else {
      if (this.storeIds != '') {
        this.expandedGLRows = [];
        this.setHeaderData()
        this.searchGLData();
        this.actionType = 'Y'
      } else {
        // this.NoData = '';
      }
    }
  }
  togglePopover(popoverIndex: number) {
    this.activePopover = this.activePopover === popoverIndex ? -1 : popoverIndex;
  }

  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    } else if (value < 0) {
      return false;
    }
    return true;
  }
  ngAfterViewInit(): void {
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.shared.common.pageName == 'GL Lookup') {
        if (res.obj.storesData != undefined) {
          this.groupsArray = res.obj.storesData;
          this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
          this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_name : this.groupName = ''
          this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
          this.getStoresandGroupsValues()
        }
      }
    })

    this.excel = this.apiSrvc.getExportToExcelAllReports().subscribe((res: { obj: { state: boolean; title: string; }; }) => {
      if (this.excel != undefined) {
        if (res.obj.title == 'GL Lookup') {
          if (res.obj.state == true) {
            this.exportAsXLSX();
          }
        }
      }
    });
    this.print = this.apiSrvc.getExportToPrintAllReports().subscribe((res: { obj: { state: boolean; title: string; }; }) => {
      if (this.print != undefined) {
        if (res.obj.title == 'GL Lookup') {
          if (res.obj.state == true) {
            // this.GetPrintData();
          }
        }
      }
    });
    this.Pdf = this.apiSrvc.getExportToPDFAllReports().subscribe((res: { obj: { state: boolean; title: string; }; }) => {
      if (this.Pdf != undefined) {
        if (res.obj.title == 'GL Lookup') {
          if (res.obj.state == true) {
            // this.generatePDF();
          }
        }
      }
    });
    this.email = this.apiSrvc.getExportToEmailPDFAllReports().subscribe((res: { obj: { state: boolean; title: string; }; }) => {
      if (this.email != undefined) {
        if (res.obj.title == 'GL Lookup') {
          if (res.obj.state == true) {
            // this.sendEmailData(res.obj.Email, res.obj.notes, res.obj.from);
          }
        }
      }
    });
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

    const formattedMonth = this.month
      ? this.shared.datePipe.transform(this.month, 'MMMM yyyy')
      : '';

    return {
      title: 'GL Lookup',
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
          value: formattedMonth || ''
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
  ExcelStoreNames: any = [];
  exportAsXLSX() {

    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet("GL Lookup");

    const DATE_EXTENSION = this.shared.datePipe.transform(new Date(), 'MMddyyyy');

    /* ===========================================================
        1️⃣ FILTER DATA BY STORE (fix for exporting all stores)
    ============================================================ */
    const exportData = [...this.GLData];

    /* ===========================================================
        2️⃣ TOP FILTER SECTION
    ============================================================ */

    worksheet.addRow(["Store:", this.getStoreNames()]);
    worksheet.getRow(1).font = { bold: true, size: 11 };

    worksheet.addRow(["Account Number:", this.accountNumber || "-"]);
    worksheet.getRow(2).font = { bold: true, size: 11 };

    const timeFrame = this.month
      ? this.month.toLocaleString("en-US", { month: "long", year: "numeric" })
      : "-";

    worksheet.addRow(["Time Frame:", timeFrame]);
    worksheet.getRow(3).font = { bold: true, size: 11 };

    worksheet.addRow([]);
    const headers = [
      "ACCOUNT NUMBER",
      "STORE",
      "CONTROL NUMBER",
      "CONTROL NUMBER 2",
      "DETAIL DESCRIPTION",
      "POSTING AMOUNT",
      "ACCOUNTING DATE"
    ];

    const headerRow = worksheet.addRow(headers);

    headerRow.font = { bold: true, size: 12, color: { argb: "FFFFFF" } };
    headerRow.alignment = { horizontal: "center", vertical: "middle" };

    headerRow.eachCell((cell) => {
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "2a91f0" }
      };
      cell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" }
      };
    });

    worksheet.views = [{ state: "frozen", ySplit: 5 }];

    /* ===========================================================
        4️⃣ COLUMN WIDTHS
    ============================================================ */

    worksheet.columns = [
      { key: "Accountnumber", width: 18 },
      { key: "Store", width: 25 },
      { key: "CountrolNumber", width: 20 },
      { key: "CountrolNumber2", width: 20 },
      { key: "detaildescription", width: 45 },
      { key: "postingamount", width: 18 },
      { key: "accountingdate", width: 18 }
    ];

    /* ===========================================================
        5️⃣ TABLE ROWS
    ============================================================ */

    let rowIndex = 6;

    exportData.forEach((item: any) => {

      const row = worksheet.addRow({
        Accountnumber: item.Accountnumber || "-",
        Store: item.Store || "-",
        CountrolNumber: item.CountrolNumber || "-",
        CountrolNumber2: item.CountrolNumber2 || "-",
        detaildescription: item.detaildescription || "-",
        postingamount: item.postingamount || 0,
        accountingdate: item.accountingdate ? new Date(item.accountingdate) : "-"
      });

      // Currency
      row.getCell("postingamount").numFmt = "$#,##0.00";

      // Date
      if (item.accountingdate) {
        row.getCell("accountingdate").numFmt = "mm/dd/yyyy";
      }

      // Borders & alignment
      row.eachCell((cell) => {
        cell.border = {
          top: { style: "thin" },
          bottom: { style: "thin" },
          left: { style: "thin" },
          right: { style: "thin" }
        };
        cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
      });

      // Zebra rows
      if (rowIndex % 2 === 0) {
        row.eachCell((cell) => {
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "e5e5e5" }
          };
        });
      }

      rowIndex++;
    });

    /* ===========================================================
        6️⃣ GRAND TOTAL ROW
    ============================================================ */

    const totalRow = worksheet.addRow([
      "", "", "", "",
      "TOTAL",
      this.grandTotal,
      ""
    ]);

    totalRow.font = { bold: true };

    totalRow.getCell(6).numFmt = "$#,##0.00";

    totalRow.eachCell((cell) => {
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "d9ead3" }
      };
      cell.border = {
        top: { style: "thin" },
        bottom: { style: "thin" },
        left: { style: "thin" },
        right: { style: "thin" }
      };
      cell.alignment = { vertical: "middle", horizontal: "center" };
    });

    /* ===========================================================
        7️⃣ DOWNLOAD EXCEL
    ============================================================ */

    workbook.xlsx.writeBuffer().then((buffer) => {
      FileSaver.saveAs(
        new Blob([buffer], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        }),
        `GLLookup_Report_${DATE_EXTENSION}.xlsx`
      );
    });
  }




  // ---------------------------  STORE NAME HELPER ---------------------------
  getStoreNames(): string {
    const allStores = this.shared.common.groupsandstores.flatMap((g: any) => g.Stores);

    // REMOVE DUPLICATES
    const uniqueIds = [...new Set(this.storeIds)];

    // Check if all stores selected
    if (uniqueIds.length === allStores.length) {
      return "All Stores";
    }

    // Correct selection mapping
    const selected = allStores
      .filter((s: any) => uniqueIds.includes(s.ID))
      .map((s: any) => s.storename);

    return selected.join(", ");
  }
}
