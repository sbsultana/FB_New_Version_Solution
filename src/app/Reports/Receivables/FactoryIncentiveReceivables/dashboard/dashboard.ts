import { Component, HostListener } from '@angular/core';
import { NgbActiveModal, NgbDateParserFormatter, NgbModal, NgbModule } from '@ng-bootstrap/ng-bootstrap';
import { Workbook } from 'exceljs';
import { DatePipe } from '@angular/common';
import html2canvas from 'html2canvas';
import jsPDF from 'jspdf';
import { Subscription } from 'rxjs';
// import { DealRecapComponent } from 'src/app/Global/cdpdataview/deal/deal-recap/deal-recap.component';

// Shared / app-level utilities & modules (adapted to your project's shared service/module)
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { Setdates } from '../../../../Core/Providers/SetDates/setdates';
import { BsDatepickerConfig, BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { FilterPipe } from '../../../../Core/Providers/filterpipe/filter.pipe';
import { ChangeDetectorRef } from '@angular/core';
import { TimeConversionPipe } from '../../../../Core/Providers/pipes/timeconversion.pipe';
import { NgxSpinnerService } from 'ngx-spinner';
import { Router } from '@angular/router';
import { common } from '../../../../common';
import { Stores } from '../../../../CommonFilters/stores/stores';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';

@Component({
  selector: 'app-dashboard',
  imports: [SharedModule, BsDatepickerModule, TimeConversionPipe, Stores, NgbModule],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
  standalone: true,
})
export class Dashboard {
  /* ------------------------------- VARIABLES ------------------------------- */

  NoData: boolean = false;
  FloorPlanData: any = [];
  FloorPlanTotalData: any = [];

  dealStatus: any = ['Booked', 'Finalized', 'Delivered'];
  dealType: any = ['Retail'];
  allordebit: any = 'dr';
  selectedreceviabe: any = [];
  StoreVal: any = '71,53,8,7,4,35,1,32,40,50,25,18,31,3,70,72,2,17,41,55,42,51,12,73,54,9,15,5,14,30,11';
  spinnerLoader: boolean = true;
  Role: any = [];
  userid: any;

  index: string = '';
  groups: any = 1;
  financeManagerId: any = '0';
  AgeFrom: any = 0;
  AgeTo: any = 0;
  QISearchName: any = '';

  callLoadingState = 'FL';
  enablevehicle: boolean = false;
  commentsVisibility: boolean = true;

  hideVisibility: boolean = false;
  hideRecords: any = [];
  FinalArray: any = [];

  header: any = [
    {
      type: 'Bar',

      storeIds: this.StoreVal,
      groups: this.groups,
      financemanagers: this.financeManagerId,
      ageFrom: this.AgeFrom,
      ageTo: this.AgeTo,
    },
  ];

  excel!: Subscription;
  Pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  reportgetting!: Subscription;

  popup: any = [{ type: 'Popup' }];
  // actionType: any = 'N';

  notesStageValue: any = '';
  notesStageText: any = '';
  notesstage: any = [];
  selecteddata: any = [];
  check: boolean = false;
  activePopover: number = -1;

  selectedFiManagersvalues: any = [];
  selectedFiManagersname: any = [];
  financeManager: any = [];
  storename: any = '';
  groupsArray: any = [];
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  groupId: any = 8;
  storeIds: any = '';
  stores: any = [];

  storesFilterData: any = {
    groupsArray: this.groupsArray,
    groupId: this.groupId,
    storesArray: this.stores,
    storeids: '1',
    type: 'M',
    others: 'N',
    groupName: this.groupName,
    storename: this.storename,
    storecount: null,
    storedisplayname: this.storedisplayname,
  };
  // solutionurl: any = environment.apiUrl;
  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent) {
    const clickedInside = (event.target as HTMLElement).closest(
      '.dropdown-toggle, .dropdown-menu , .timeframe, .reportstores-card',
    );
    if (!clickedInside) {
      this.activePopover = -1;
    }
  }

  /* ------------------------------- CONSTRUCTOR ------------------------------- */

  constructor(
    private ngbmodal: NgbModal,
    private comm: common,
    public shared: Sharedservice,
    private router: Router,
    private ngbmodalActive: NgbActiveModal,
    private spinner: NgxSpinnerService,
    private toast: ToastService,
    private datepipe: DatePipe,
  ) {
    if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
      if (localStorage.getItem('flag') == 'V') {
        this.storeIds = [];
        console.log(JSON.parse(localStorage.getItem('userInfo')!), JSON.parse(localStorage.getItem('userInfo')!).user_Info, 'Widget Stores............');
        this.groupId = JSON.parse(localStorage.getItem('userInfo')!).groupid
        JSON.parse(localStorage.getItem('userInfo')!).store.indexOf(',') > 0 ?
          this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).store.split(',') :
          this.storeIds.push(JSON.parse(localStorage.getItem('userInfo')!).store)
        this.financeManagerId = 0;
        JSON.parse(localStorage.getItem('userInfo')!).flag2 == 'A' ? this.allordebit = 'all' : this.allordebit = 'dr';
        localStorage.setItem('flag', 'M')
      } else {
        this.groupId = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Preferences
        this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids.split(',')
      }
      this.Role = JSON.parse(localStorage.getItem('userInfo')!).user_Info.title;
      this.userid = JSON.parse(localStorage.getItem('userInfo')!).user_Info.userid;
    }


    this.commentsVisibility = true;
    this.shared.setTitle(this.comm.titleName + '-Factory Incentive Receivables');

    /* --------- HEADER FOR REPORTING --------- */
    const data = {
      title: 'Factory Incentive Receivables',
      stores: this.StoreVal,
      groups: this.groups,
      financemanagers: '',
      count: 0,
      AgeFrom: this.AgeFrom,
      AgeTo: this.AgeTo,
      search: this.QISearchName,
    };

    this.shared.api.SetHeaderData({ obj: data });

    this.header = [
      {
        type: 'Bar',
        storeIds: this.StoreVal,
        groups: this.groups,
        financemanagers: this.financeManagerId,
        ageFrom: this.AgeFrom,
        ageTo: this.AgeTo,
      },
    ];

    /* --------- AUTO LOAD FIRST TAB --------- */
    if (this.StoreVal != '') {
      this.Getfloorplansdata(this.selectedreceviabe);
      this.getEmployees();
    }
  }

  /* ------------------------------- LIFECYCLE ------------------------------- */

  ngOnInit(): void { }

  ngAfterViewInit(): void {
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.shared.common.pageName == 'Factory Incentive Receivables') {
        if (res.obj.storesData != undefined) {
          this.groupsArray = res.obj.storesData;
          this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
          this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_name : this.groupName = ''
          this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
          this.getStoresandGroupsValues()
        }
      }
    })
    this.shared.api.GetReportOpening().subscribe((res) => {
      if (res.obj.Module === 'Factory Incentive Receivables') {
        document.getElementById('report')?.click();
      }
    });

    /* -------- REFRESH DATA FROM HEADER -------- */
    this.reportgetting = this.shared.api.GetReports().subscribe((data) => {
      if (data.obj.Reference === 'Factory Incentive Receivables') {
        this.FloorPlanData = [];
        // this.actionType = 'Y';
        this.NoData = false;

        /* Update filters */
        if (!data.obj.header) {
          this.StoreVal = data.obj.storeValues;
          this.financeManagerId = data.obj.FIvalues;
          this.groups = data.obj.groups;
          this.AgeFrom = data.obj.AgeFrom;
          this.AgeTo = data.obj.AgeTo;
        }

        if (this.StoreVal != '') {
          this.Getfloorplansdata(this.selectedreceviabe);
        } else {
          this.NoData = true;
        }

        /* Update header for next reload */
        const headerdata = {
          title: 'Factory Incentive Receivables',
          stores: this.StoreVal,
          groups: this.groups,
          financemanagers: this.financeManagerId,
          AgeFrom: this.AgeFrom,
          AgeTo: this.AgeTo,
        };

        this.shared.api.SetHeaderData({ obj: headerdata });

        this.header = [
          {
            type: 'Bar',
            storeIds: this.StoreVal,
            groups: this.groups,
            financemanagers: this.financeManagerId,
            ageFrom: this.AgeFrom,
            ageTo: this.AgeTo,
          },
        ];
      }
    });

    this.excel = this.shared.api.getExportToExcelAllReports().subscribe((res: any) => {
      if (res && res.obj && res.obj.title == 'Factory Incentive Receivables' && res.obj.state == true) {
        this.exportToExcel(); // merged export will create both sheets
      }
    });
  }
  //   formatBalance(val: number | null, decimals = 2): string {
  //   if (val === null || val === undefined) return '-';
  //   return val < 0
  //     ? `-$${Math.abs(val).toLocaleString(undefined, { minimumFractionDigits: decimals })}`
  //     : `$${val.toLocaleString(undefined, { minimumFractionDigits: decimals })}`;
  // }

  formatBalance(val: number | null): string {
    if (val === null || val === undefined) return '-';
    return val < 0 ? `-$${Math.abs(val).toFixed(2)}` : `$${val.toFixed(2)}`;
  }

  formatBalancetozero(val: number | null): string {
    if (val === null || val === undefined) return '-';
    if (val === 0) return '0';
    return val < 0 ? `-$${Math.abs(val).toFixed(0)}` : `$${val.toFixed(0)}`;
  }
  exportToExcel(): void {
    const workbook: any = this.shared.getWorkbook?.();
    if (!workbook) {
      this.toast.show('Workbook helper not available', 'danger', 'Error');
      return;
    }

    /* ================= SAFE NAME HELPERS ================= */

    const getSafeSheetName = (name: string): string => {
      const cleaned = (name || 'Sheet1')
        .replace(/[\\/*?:[\]]/g, '')
        .substring(0, 31)
        .trim();

      return cleaned || 'Sheet1';
    };

    const getSafeFileName = (name: string): string => {
      return (name || 'Report')
        .replace(/[\\/*?:[\]]/g, '')
        .trim();
    };

    /* ================= REPORT TITLE ================= */

    const reportTitle =
      this.selectedreceviabe?.title
        ? `Factory Incentive Receivables`
        : 'Factory Incentive Receivables';

    const safeSheetName = getSafeSheetName(reportTitle);

    // Prevent duplicate sheet names
    let finalSheetName = safeSheetName;
    let counter = 1;
    while (workbook.getWorksheet(finalSheetName)) {
      finalSheetName = `${safeSheetName}_${counter++}`;
    }

    const summarySheet = workbook.addWorksheet(finalSheetName);

    /* ================= TITLE ================= */

    const titleRow = summarySheet.addRow([reportTitle]);
    titleRow.font = { bold: true, size: 14 };

    /* ================= FILTERS ================= */

    summarySheet.addRow([]);

    /* ================= AGING HEADER ================= */

    const agingHeader = summarySheet.addRow(['', 'Total', '0-5', '6-10', '11-15', '15+']);
    agingHeader.font = { bold: true, color: { argb: 'FFFFFFFF' } };

    agingHeader.eachCell((cell: any) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2A91F0' } };
      cell.alignment = { horizontal: 'center' };
    });

    if (this.FloorPlanData?.length) {
      const A = this.FloorPlanData[0].AgeData?.[0];

      const agingValueRow = summarySheet.addRow([
        `Vehicle Aging`,
        A?.TOTAL ?? null,
        A?.D1 ?? null,
        A?.D2 ?? null,
        A?.D3 ?? null,
        A?.D4 ?? null,
      ]);

      [2, 3, 4, 5, 6].forEach((col) => {
        const cell = agingValueRow.getCell(col);
        if (typeof cell.value === 'number') {
          cell.numFmt = '$#,##0.00;[Red]-$#,##0.00';
          cell.alignment = { horizontal: 'right' };
        }
      });

      agingValueRow.font = { bold: true };
    }

    summarySheet.addRow([]);

    /* ================= DETAIL HEADERS ================= */

    const headers = [
      'Age', 'Date', 'Account', 'Control', 'Control 2', 'Balance', 'Customer',
      'Number', 'Sale Date', 'Sale Age', 'Stock #', 'Deal #', 'Stage',
      'F&I Mgr', 'New/Used', 'Type', 'Status', 'Bank Name',
      'Year', 'Make', 'Model',
    ];

    const headerRow = summarySheet.addRow(headers);
    headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };

    headerRow.eachCell((cell: any) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2A91F0' } };
      cell.alignment = { horizontal: 'center' };
    });

    /* ================= DATE FORMATTER ================= */

    const formatDate = (d: any) => {
      if (!d) return '-';
      const dt = new Date(d);
      if (isNaN(dt.getTime())) return '-';
      return (
        `${String(dt.getMonth() + 1).padStart(2, '0')}.` +
        `${String(dt.getDate()).padStart(2, '0')}.` +
        `${dt.getFullYear()}`
      );
    };

    /* ================= DATA ROWS ================= */

    (this.FloorPlanData || []).forEach((fp: any) => {
      const row = summarySheet.addRow([
        fp.AGE || '-',
        formatDate(fp.FundedDate),
        fp.AccountDesc2 || '-',
        fp.Control || '-',
        fp.control2 || '-',
        fp.Balance ?? null,
        fp.CustomerName || '-',
        fp.CustomerNumber || '-',
        formatDate(fp.SaleDate),
        fp.SaleAge || '-',
        fp.StockNo || '-',
        fp.DealNo || '-',
        fp.Stage || '-',
        fp.FIManager || '-',
        fp.DealType || '-',
        fp.SaleType || '-',
        fp.DealStatus || '-',
        fp.BankName || '-',
        fp.VehicleYear || '-',
        fp.VehicleMake || '-',
        fp.VehicleModel || '-',
      ]);

      const balanceCell = row.getCell(6);
      if (typeof balanceCell.value === 'number') {
        balanceCell.numFmt = '$#,##0.00;[Red]-$#,##0.00';
        balanceCell.alignment = { horizontal: 'right' };
      }

      if (fp.duplicateNotes?.length) {
        summarySheet.addRow([]);
        const notesHeader = summarySheet.addRow(['Notes']);
        notesHeader.font = { bold: true };

        fp.duplicateNotes.forEach((n: any) => {
          summarySheet.addRow([n.NOTES || '-']);
        });
      }
    });

    summarySheet.columns.forEach((c: any) => (c.width = 20));

    /* ================= TOTALS ROW ================= */

    if (this.FloorPlanTotalData?.length > 0) {
      const totalBalance = this.FloorPlanTotalData[0].Balance ?? 0;

      summarySheet.addRow([]);

      const totalsRow = summarySheet.addRow([
        'Totals', '', '', '', '',
        totalBalance,
        '', '', '', '', '', '', '', '', '', '', '', '', '', '',
      ]);

      totalsRow.font = { bold: true };
      summarySheet.mergeCells(`A${totalsRow.number}:E${totalsRow.number}`);
      totalsRow.getCell(1).alignment = { horizontal: 'right' };

      const balanceCell = totalsRow.getCell(6);
      balanceCell.numFmt = '$#,##0.00;[Red]-$#,##0.00';
      balanceCell.alignment = { horizontal: 'right' };
    }

    /* ================= EXPORT ================= */

    workbook.xlsx.writeBuffer().then(() => {
      const rawFileName =
        `${'Factory Incentive Receivables'}_Report`;

      const safeFileName = getSafeFileName(rawFileName);

      this.shared.exportToExcel(workbook, safeFileName);
    });
  }

  keyPressNumbers(event: any) {
    var charCode = event.which ? event.which : event.keyCode;
    // Only Numbers 0-9
    if (charCode < 48 || charCode > 57) {
      event.preventDefault();
      return false;
    } else {
      return true;
    }
  }

  /* ------------------------------- SEARCH HANDLER ------------------------------- */

  receiveMessage($event: any) {
    this.QISearchName = $event;
  }

  /* ------------------------------- SORT ------------------------------- */

  isDesc: boolean = false;
  column: string = 'CategoryName';

  // sort(property: any, state?: any) {
  //   if (state === undefined) {
  //     this.isDesc = !this.isDesc;
  //   }

  //   this.callLoadingState = 'FL';
  //   this.column = property;

  //   let direction = this.isDesc ? 1 : -1;

  //   this.FloorPlanData.sort((a: any, b: any) => {
  //     if (a[property] < b[property]) return -1 * direction;
  //     if (a[property] > b[property]) return 1 * direction;
  //     return 0;
  //   });
  // }

  sort(property: string, state?: any) {
    if (state === undefined) {
      this.isDesc = !this.isDesc;
    }

    this.callLoadingState = 'FL';
    this.column = property;

    const direction = this.isDesc ? 1 : -1;

    this.FloorPlanData.sort((a: any, b: any) => {
      let valA = a[property];
      let valB = b[property];

      // Normalize null / dash
      if (valA === null || valA === undefined || valA === '-') valA = '';
      if (valB === null || valB === undefined || valB === '-') valB = '';

      // 🔑 Detect numeric values (Control, Deal #, Stock # when numeric)
      const numA = Number(valA);
      const numB = Number(valB);

      const isNumA = !isNaN(numA);
      const isNumB = !isNaN(numB);

      // ✅ Numeric comparison when both are numbers
      if (isNumA && isNumB) {
        return (numA - numB) * direction;
      }

      // ✅ String comparison fallback
      return String(valA).localeCompare(String(valB), undefined, {
        numeric: true,
        sensitivity: 'base'
      }) * direction;
    });
  }

  /* ------------------------------- MAIN API CALL ------------------------------- */
  previousReportPath: string | null = null;
  Getfloorplansdata(path: any) {
    this.goToFirstPage();
    if (this.previousReportPath !== path.path) {
      this.AgeFrom = 0;
      this.AgeTo = 0;
      this.previousReportPath = path.path;
    }
    this.selectedreceviabe = path;
    this.NoData = false;
    this.FloorPlanData = [];
    this.FloorPlanTotalData = [];
    this.filteredFloorplanData = [];
    if (
      this.AgeFrom !== null &&
      this.AgeTo !== null &&
      Number(this.AgeFrom) > Number(this.AgeTo)
    ) {
      this.toast.show('Please Enter Valid Age Range', 'warning', 'Warning');
      return; // ⛔ stop execution
    }
    this.spinner.show();


    console.log('hi')

    const obj = {
      AS_ID: this.storeIds.toString(),
      UserID: 0,
      FIR_Type: this.dealType.toString(),
      ValueType: this.allordebit.toString() == 'all' ? '' : this.allordebit.toString(),
    };

    let startFrom = new Date().getTime();

    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetFactoryIncentiveReceivables', obj).subscribe(
      (res) => {
        if (res.status == 200 && res.response) {
          this.spinner.hide();

          if (res.response.length > 0) {
            this.FloorPlanData = res.response.filter((e: any) => e.store !== 'TOTAL');
            this.FloorPlanTotalData = res.response.filter((e: any) => e.store === 'TOTAL');
            this.FloorPlanData.forEach((x: any) => {
              x.AgeData = JSON.parse(x.AgeData);

              if (x.Comments) x.Comments = JSON.parse(x.Comments);
              if (x.Notes) {
                x.Notes = JSON.parse(x.Notes);
                x.duplicateNotes = [...x.Notes];
                x.Notesstate = '+';

                if (x.Notes.length > 3) {
                  x.duplicateNotes = x.duplicateNotes.slice(0, 3);
                }
              }
            });

            if (this.callLoadingState == 'ANS') {
              this.sort(this.column, this.callLoadingState);
            }
            this.filteredFloorplanData = this.FloorPlanData || [];
            this.NoData = this.FloorPlanData.length == 0;
          } else {
            this.NoData = true;
          }
        } else {
          this.NoData = true;
          this.spinner.hide();
        }
      },
      () => {
        this.toast.show('502 Bad Gateway Error', 'danger', 'Error');
        this.spinner.hide();
        this.NoData = true;
      },
    );
  }

  ////////////////////////////////////////////Pagination Code////////////////////////////////////

  searchText: string = '';
  currentPage: number = 1;
  itemsPerPage: number = 100;
  maxPageButtonsToShow: number = 5;
  clickedPage: number | null = null;

  filteredFloorplanData: any[] = [];
  setFloorplanData() {
    this.filteredFloorplanData = this.FloorPlanData || [];
    this.currentPage = 1; // reset page
  }
  get filteredData() {

    if (!this.searchText) {
      return this.filteredFloorplanData;
    }

    const searchTerms = this.searchText
      .split(',')
      .map(term => term.trim().toLowerCase())
      .filter(term => term.length > 0);

    const filtered = this.filteredFloorplanData.filter((x: any) =>
      searchTerms.some(term =>

        x.AGE?.toString().toLowerCase().includes(term) ||
        x.FundedDate?.toString().toLowerCase().includes(term) ||
        x.AccountDesc2?.toLowerCase().includes(term) ||

        x.Control?.toLowerCase().includes(term) ||
        x.control2?.toLowerCase().includes(term) ||

        x.Balance?.toString().toLowerCase().includes(term) ||

        x.CustomerName?.toLowerCase().includes(term) ||
        x.CustomerNumber?.toString().toLowerCase().includes(term) ||

        x.SaleDate?.toString().toLowerCase().includes(term) ||
        x.SaleAge?.toString().toLowerCase().includes(term) ||

        x.store?.toLowerCase().includes(term) ||

        x.StockNo?.toString().toLowerCase().includes(term) ||
        x.DealNo?.toString().toLowerCase().includes(term) ||

        x.Stage?.toLowerCase().includes(term) ||
        x.FIManager?.toLowerCase().includes(term) ||

        x.DealType?.toLowerCase().includes(term) ||
        x.SaleType?.toLowerCase().includes(term) ||
        x.DealStatus?.toLowerCase().includes(term) ||

        x.BankName?.toLowerCase().includes(term) ||

        x.VehicleYear?.toString().toLowerCase().includes(term) ||
        x.VehicleMake?.toLowerCase().includes(term) ||
        x.VehicleModel?.toLowerCase().includes(term)

      )
    );
    return filtered.sort((a: any, b: any) => {
      const aIndex = searchTerms.findIndex(term =>
        Object.values(a).join(' ').toLowerCase().includes(term)
      );
      const bIndex = searchTerms.findIndex(term =>
        Object.values(b).join(' ').toLowerCase().includes(term)
      );
      return aIndex - bIndex;
    });
  }
  sortColumn: string = '';
  sortDirection: 'asc' | 'desc' = 'asc';
  sortTable(column: string) {
    if (this.sortColumn === column) {
      this.sortDirection = this.sortDirection === 'asc' ? 'desc' : 'asc';
    } else {
      this.sortColumn = column;
      this.sortDirection = 'asc';
    }
    this.filteredFloorplanData.sort((a: any, b: any) => {
      let valA = a[column];
      let valB = b[column];
      valA = valA ?? '';
      valB = valB ?? '';
      if (!isNaN(valA) && !isNaN(valB)) {
        return this.sortDirection === 'asc'
          ? valA - valB
          : valB - valA;
      }
      return this.sortDirection === 'asc'
        ? valA.toString().localeCompare(valB.toString())
        : valB.toString().localeCompare(valA.toString());

    });

  }
  getSortIcon(column: string) {
    if (this.sortColumn !== column) return 'fa-sort';
    return this.sortDirection === 'asc'
      ? 'fa-sort-up'
      : 'fa-sort-down';
  }
  get paginatedItems() {
    const startIndex = (this.currentPage - 1) * this.itemsPerPage;
    const endIndex = startIndex + this.itemsPerPage;
    return this.filteredData.slice(startIndex, endIndex);
  }
  onPageSizeChange() {
    this.currentPage = 1;
  }
  getMaxPageNumber(): number {
    return Math.ceil(this.filteredData.length / this.itemsPerPage);
  }
  nextPage() {
    if (this.currentPage < this.getMaxPageNumber()) this.currentPage++;
  }
  prevPage() {
    if (this.currentPage > 1) this.currentPage--;
  }
  goToFirstPage() {
    this.currentPage = 1;
  }

  goToLastPage() {
    this.currentPage = this.getMaxPageNumber();
  }
  getStartRecordIndex(): number {
    return (this.currentPage - 1) * this.itemsPerPage;
  }
  getEndRecordIndex(): number {
    const endIndex = this.getStartRecordIndex() + this.itemsPerPage;
    return endIndex > this.filteredData.length
      ? this.filteredData.length
      : endIndex;
  }

  get BalanceTotal(): number {
    return this.filteredData.reduce((total, item) => {
      return total + (parseFloat(item.Balance) || 0);
    }, 0);
  }

  get BalFloorplanTotal(): number {
    return this.filteredData.reduce((total, item) => {
      return total + (parseFloat(item.BalFloorplan) || 0);
    }, 0);
  }
  ////////////////////////////////////Pagination Code End////////////////////////////////////////////




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
  AllorDebit(e: any) {
    this.allordebit = []
    // if (e == 'All') {
    //   if (this.dealStatus.length == 3) {
    //     this.dealStatus = []
    //     this.toast.warning('Please Select Atleast One dealStatus', '');

    //   } else {
    //     this.dealStatus = []
    //     this.dealStatus = ['Booked', 'Finalized','Delivered'];
    //   }
    // } 
    // else {
    const index = this.allordebit.findIndex((i: any) => i == e);
    if (index >= 0) {
      this.allordebit.splice(index, 1);
    } else {
      this.allordebit.push(e);
    }
    // }
  }


  multipleorsingle(ref: any, e: any) {
    if (ref == 'DS') {
      const index = this.dealStatus.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.dealStatus.splice(index, 1);
      } else {
        this.dealStatus.push(e);
      }
    }
    if (ref == 'DT') {
      const index = this.dealType.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.dealType.splice(index, 1);
      } else {
        this.dealType.push(e);
      }
    }
  }

  /* ------------------------------- NOTES LOGIC ------------------------------- */

  viewmoreAction(fp: any) {
    if (fp.Notesstate === '+') {
      fp.Notesstate = '-';
      fp.duplicateNotes = fp.Notes;
    } else {
      fp.Notesstate = '+';
      fp.duplicateNotes = fp.Notes.slice(0, 3);
    }
  }

  getDropDown(companyid: any) {
    const obj = {
      AssociatedReport: this.selectedreceviabe.path,
      CompanyID: companyid,
    };

    this.shared.api
      .postmethod(this.comm.routeEndpoint + 'GetScheduleNoteStages', obj)
      .subscribe((res) => {
        if (res.status == 200) {
          this.notesstage = res.response;
          this.notesStageValue = '';
        }
      });
  }

  addNotes(item: any) {
    this.selecteddata = item;
    this.getDropDown(item.companyid);
    this.notesStageText = '';
    this.notesStageValue = '';
  }

  formatDate(date?: number | string | Date): string {
    if (!date) return '-';

    const parsedDate = new Date(date);

    if (isNaN(parsedDate.getTime())) return '-';

    return new Intl.DateTimeFormat('en-US', {
      month: '2-digit',
      day: '2-digit',
      year: 'numeric',
    }).format(parsedDate);
  }
  save() {
    if (this.notesStageText.trim() === '') {
      this.toast.show('Please enter notes', 'warning', 'Warning');
      return;
    }

    const obj = {
      AS_ID: this.selecteddata.storeid,
      Account: this.selecteddata.Account,
      Control: this.selecteddata.Control,
      Notes: this.notesStageText,
      StageId: this.notesStageValue,
      UserID: this.userid,
    };

    this.shared.api
      .postmethod(this.comm.routeEndpoint + 'AddScheduleNotesAction', obj)
      .subscribe((res: any) => {
        if (res.status == 200) {
          this.toast.show('Notes Added Successfully', 'success', 'Success');
          this.callLoadingState = 'ANS';
          (document.getElementById('close') as HTMLInputElement)?.click();
          this.oncloseone();
          this.commentsVisibility = true;

          /* Instant Frontend Update */
          const userName = JSON.parse(localStorage.getItem('UserDetails')!).UserName;
          const curDate = new Date();
          let nts = '';

          if (this.notesStageValue) {
            const filtered = this.notesstage.find(
              (item: any) => item.NS_ID == this.notesStageValue,
            );
            nts = `[${filtered.NS_Text}] ${this.notesStageText} - ${userName} - ${this.formatDate(curDate)}`;
            this.selecteddata.Stage = filtered.NS_Text;
          } else {
            nts = `${this.notesStageText} - ${userName} - ${this.formatDate(curDate)}`;
          }

          const newNote = {
            STAGE: '',
            NOTES: nts,
            NOTESDATE: this.formatDate(curDate),
            UserName: userName,
          };

          if (!this.selecteddata.duplicateNotes) {
            this.selecteddata.duplicateNotes = [];
          }

          this.selecteddata.duplicateNotes.unshift(newNote);
          this.selecteddata.NotesStatus = 'Y';
        } else {
          this.toast.show('Something went wrong. Please try again.', 'danger', 'Error');
        }
      });
  }

  /* ------------------------------- HIDE RECORD LOGIC ------------------------------- */

  collectHidevalues(e: any, val: any, confirmtemplate: any, ref: any, refval: any) {
    if (ref === 'multi') {
      if (this.hideRecords.length === 0) {
        this.toast.show('Please Select Atleast One Record to Hide', 'warning', 'Warning');
        (document.getElementById('symbol') as HTMLInputElement).checked = false;
        return;
      }

      if (e.target.checked) {
        this.hideVisibility = true;
        this.ngbmodalActive = this.ngbmodal.open(confirmtemplate, {
          size: 'sm',
          backdrop: 'static',
        });
      }
    } else {
      if (e.target.checked) {
        this.hideVisibility = true;
        this.hideRecords.push(val);
      } else {
        const index = this.hideRecords.findIndex((list: any) => list.StockNo == refval);
        this.hideRecords.splice(index, 1);
      }
    }
  }

  oncloseone() {
    (document.getElementById('symbol') as HTMLInputElement).checked = false;
    this.ngbmodalActive.close();
  }

  hideAdd() {
    if (this.hideRecords.length === 0) {
      this.toast.show('Please Select Atleast One Record to Hide', 'warning', 'Warning');
      return;
    }

    this.FinalArray = this.hideRecords.map((item: any) => ({
      Receivable_Type: this.selectedreceviabe.path,
      Account: item.Account,
      CompanyID: item.companyid,
      AS_ID: item.storeid,
      Control: item.Control,
      Stock: item.StockNo,
      Control_Status: 'Y',
      Deal: item.DealNo,
      UserID: this.userid,
    }));

    const obj = { receivableexcludecontrol: this.FinalArray };

    this.shared.api.postmethod('ReceivableExcludeControls', obj).subscribe((res) => {
      if (res.status == 200) {
        this.toast.show('This Control Hidden Successfully', 'success', 'Success');
        (document.getElementById('closeone') as HTMLElement).click();
        this.oncloseone();
        this.Getfloorplansdata(this.selectedreceviabe);
        this.hideRecords = [];
        this.hideVisibility = false;
      } else {
        this.toast.show('Failed to hide control.', 'danger', 'Error');
      }
    });
  }

  /* ------------------------------- UTIL ------------------------------- */

  public inTheGreen(value: number): boolean {
    return value >= 0;
  }
  onclose() {
    const element = <HTMLInputElement>document.getElementById('symbol');
    if (element) element.checked = false;
    this.ngbmodalActive.close();
  }

  onclosealert() {
    const element = <HTMLInputElement>document.getElementById('symbol');
    if (element) element.checked = false;
    this.ngbmodalActive.close();
    (document.getElementById('close') as HTMLInputElement)?.click();
  }
  viewDeal(dealData: any) {
    // const modalRef = this.ngbmodal.open(DealRecapComponent, {
    //   size: 'md',
    //   windowClass: 'compModal'
    // });
    // modalRef.componentInstance.data = {
    //   dealno: dealData.DealNo,
    //   storeid: dealData.storeid,
    //   stock: dealData.StockNo,
    //   vin: dealData.vin,
    //   custno: dealData.CustomerNumber
    // };
  }

  openComments() {
    this.commentsVisibility = !this.commentsVisibility;
  }

  togglePopover(popoverIndex: number) {
    if (this.activePopover === popoverIndex) {
      // If the same popover is clicked, close it
      this.activePopover = -1;
    } else {
      // Open the selected popover and close others
      this.activePopover = popoverIndex;
    }
  }
  viewreport() {
    this.activePopover = -1;
    let missing = false;
    if (!this.storeIds || this.storeIds.length === 0) {
      this.toast.show(
        'Please Select Atleast One Store',
        'warning',
        'Warning'
      );
      return;
    }
    if (!this.dealStatus || this.dealStatus.length === 0) {
      this.toast.show('Please Select Atleast One Deal Status.', 'warning', 'Warning');
      missing = true;
    }
    if (missing) {
      return;
    }
    this.financeManagerId =
      this.selectedFiManagersvalues.length === this.financeManager.length
        ? '0'
        : this.selectedFiManagersvalues.toString();
    this.AgeFrom = this.AgeFrom;
    this.AgeTo = this.AgeTo;
    const payload = {
      DealStatus: this.dealStatus.toString(),
      DealType2: this.dealType.toString(),
      StoreIds: this.storeIds.toString(),
      FinanceManagerId: this.financeManagerId,
    };
    this.Getfloorplansdata(this.selectedreceviabe);
  }

  getEmployees(val?: any, ids?: any, count?: any, bar?: any) {
    const obj = {
      AS_ID: this.StoreVal,
      type: 'F',
    };
    this.shared.api.postmethod(this.shared.common.routeEndpoint + 'GetEmployeesDev', obj).subscribe(
      (res: any) => {
        if (res && res.status == 200) {
          // if (val == 'F') {
          this.financeManager = res.response.filter((e: any) => e.FiName != 'Unknown');
          this.selectedFiManagersvalues = this.financeManager.map(function (a: any) {
            return a.FiId;
          });

          // if (bar == 'Bar') {
          //   if (this.employeeschanges != '') {
          //     let fiids = (this.employeeschanges || '').toString().split(',');
          //     this.selectedFiManagersvalues = fiids;
          //   }
          //   if (this.employeeschanges == '0' || this.employeeschanges == 0) {
          //     this.selectedFiManagersvalues = this.financeManager.map(function (a: any) { return a.FiId; });
          //   }
          //   if (this.employeeschanges == '') {
          //     this.selectedFiManagersvalues = [];
          //   }
          // }
          // }
        } else {
          this.toast.show('Invalid Details');
        }
      },
      (error: any) => {
        /* ignore console errors */
      },
    );
  }
  employees(block: any, e: any, ename?: any) {
    // ========== SINGLE FM TOGGLE ========== //
    if (block === 'FM') {
      const index = this.selectedFiManagersvalues.findIndex((i: any) => i == e);

      if (index >= 0) {
        // already selected -> remove
        this.selectedFiManagersvalues.splice(index, 1);
      } else {
        // not selected -> add
        this.selectedFiManagersvalues.push(e);
      }

      // Optional: set last clicked name (if you use it in UI)
      const index1 = this.selectedFiManagersvalues.findIndex((i: any) => i == e);
      if (index1 >= 0) {
        this.selectedFiManagersname = ename;
      }

      return;
    }

    // ========== SELECT ALL / CLEAR ALL ========== //
    if (block === 'AllFM') {
      // e == 0 → Select All
      // e == 1 → Clear All

      if (e === 0) {
        // SELECT ALL
        this.selectedFiManagersvalues = this.financeManager.map((fm: any) => fm.FiId);
      } else if (e === 1) {
        // CLEAR ALL
        this.selectedFiManagersvalues = [];
      }

      return;
    }
  }
}
