import {
  Component,
  OnInit,
  ViewChild,
  TemplateRef,
  ChangeDetectorRef,
  HostListener,
  Directive,
} from '@angular/core';
import { BsDatepickerConfig, BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { DatePipe, CommonModule } from '@angular/common';
import { NgbActiveModal, NgbModal, NgbModalRef } from '@ng-bootstrap/ng-bootstrap';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { Setdates } from '../../../../Core/Providers/SetDates/setdates';
import { common } from '../../../../common';

import { HttpClient } from '@angular/common/http';
import { Workbook } from 'exceljs';
import * as FileSaver from 'file-saver';
import { Subscription } from 'rxjs';
declare var bootstrap: any;

@Component({
  selector: 'app-dashboard',
  imports: [CommonModule, SharedModule, BsDatepickerModule],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
})
export class Dashboard implements OnInit {
  modalInstance: any;
  @ViewChild('detailModal') detailModal!: TemplateRef<any>;
  item: any;
  activePopover: number = -1;
  @HostListener('document:click', ['$event'])
  clickOutside(event: MouseEvent) {
    const target = event.target as HTMLElement;

    if (!target.closest('.filter-dropdowns')) {
      this.activePopover = -1;
    }
  }

  DateType: any;
  maxToMonth!: Date;
  minToMonth!: Date;
  storeIds: any[] = [];
  stores: any[] = [];
  groupsArray: any[] = [];
  groups: any[] = [];
  groupName = '';
  storeName = '';
  storesList: any[] = [];
  allStoresData: any[] = [];
  selectedVars: any = { var1: 'All' };
  isPopoverOpen = false;
  storeButtonLabel: string = 'All';
  FromDate!: string;
  ToDate!: string;
  tomonth: Date | null = null;
  previous: Date = new Date();
  previousmultiplemax: Date = new Date();
  toprevious: any;
  fromDate: any;
  toDate: any;
  filteredMonths: string[] = [];

  // NEW / integrated fields (Client A functionality)
  // table data (original)
  searchTerm: string = '';
  isMapped: boolean | null = null;
  filteredTableData: any[] = [];
  tableData: any[] = [];
  selectedCategory: string = 'Sale';
  apiUrl = 'https://tcxtractapi.axelone.app/api/WesternAuto/GetGLAccountsInfo';

  // Account mapping specific
  AccountingMapping: any[] = []; // used in integrated UI
  Pagination: boolean = false;
  PageCount: number = 1;
  LastCount: boolean = false;
  NoData: boolean = false;
  NoDatainitail: boolean = false;

  // bulk edit / modal
  bulkEditRecords: any[] = [];
  Bulk: boolean = false;
  ignorebtn: boolean = false;

  // form fields (ngModel bind)
  actbal: string = '';
  accounttype: any = 0;
  selectedaccounttypedetail: any = 0;
  departmenttype: any = 0;
  department: any = 0;
  subtype: any = 0;
  subtypedetail: any = 0;
  subtypefulldetail: any = 0;
  flexcol1: any = 0;
  financialsummary: any = 0;
  submitted: boolean = false;

  Refchange: string = '';
  defaultresponse: any = {};
  BIref: any = '';

  Accounttype: any[] = [];
  Accounttypedetail: any[] = [];
  Departmenttype: any[] = [];
  Department: any[] = [];
  Subtype: any[] = [];
  SubtypeDetail: any[] = [];
  SubtypeFullDetail: any[] = [];
  Flexcol1: any[] = [];
  Financialsummary: any[] = [];
  Balancesummary: any[] = [];

  missingFields: string[] = [];

  // modal reference if using NgbModal programmatically
  private modalRef: NgbModalRef | null = null;

  // ---------------- Header Tabs (Client A Logic) ----------------
  selectedheadertab: any[] = ['S']; // default "Sale" selected

  // ---------------- Mapped / Unmapped & Search (Client A) ----------------
  searchText: string = '';
  statusFilter: any[] = ['M']; // Default shows both
  @ViewChild('Alert') Alert!: TemplateRef<any>;

  excel!: Subscription;

  constructor(
    public shared: Sharedservice,
    public setdates: Setdates,
    private comm: common,
    private cdr: ChangeDetectorRef,
    private datepipe: DatePipe,
    public modalService: NgbModal,
    private http: HttpClient,
     private ngbmodalActive: NgbActiveModal,
  ) {
    this.shared.setTitle(this.shared.common.titleName + '-AccountMapping');
    if (typeof window !== 'undefined') {
      if (localStorage.getItem('UserDetails') != null) {
        // existing bootstrap logic kept
      }

      this.shared.setTitle(this.shared.common.titleName + '-AccountMapping');
      if (localStorage.getItem('Fav') != 'Y') {
        const data = {
          title: 'Account Mapping',
          stores: this.storeIds,
          datetype: 'MTD',
          fromdate: this.FromDate,
          todate: this.ToDate,
        };
        this.shared.api.SetHeaderData({
          obj: data,
        });
      }
    }
  }

  getMonthYear(): string {
    if (!this.frommonth) return '';

    const dateObj = this.frommonth instanceof Date ? this.frommonth : new Date(this.frommonth);
    const month = (dateObj.getMonth() + 1).toString().padStart(2, '0');
    const year = dateObj.getFullYear();

    return `${month}-${year}`;
  }

  togglePopover(popoverId: number) {
    this.activePopover = this.activePopover === popoverId ? -1 : popoverId;
  }
  ngOnInit(): void {
    this.shared.setTitle('Account Mapping');
    this.getStoresList();

    const today = new Date();
    const fromMonth = new Date(today.getFullYear(), today.getMonth(), 1);
    const toMonth = new Date(today.getFullYear(), today.getMonth() + 1, 0);

    this.frommonth = fromMonth;
    this.tomonth = toMonth;

    this.minToMonth = this.frommonth;
    this.maxToMonth = today;
    this.getGLAccountsInfo();
  }
  onApply(): void {
    this.getGLAccountsInfo(this.searchTerm);
  }
  getStoresList(): void {
    this.shared.spinner.show();
    const obj = {
      userid: 1,
    };
    this.shared.api.postmethod('Western/GetStoresList', obj).subscribe({
      next: (res: any) => {
        console.log(res);
        if (res) {
          this.stores = res.response;
          this.storeIds = this.stores.map((s) => s.ID);

          // this.onApply();
          this.shared.spinner.hide();
        }
      },
      error: (err: any) => {
        this.shared.spinner.hide();
        console.error('Error fetching stores:', err);
      },
    });
  }

  headerselection(type: string) {
    this.selectedheadertab = [type];

    // Change category in your existing structure
    switch (type) {
      case 'S':
        this.selectedCategory = 'Sale';
        break;
      case 'C':
        this.selectedCategory = 'Cost';
        break;
      case 'X':
        this.selectedCategory = 'Non-Operating Expense';
        break;
      case 'I':
        this.selectedCategory = 'Non-Operating Income';
        break;
      case 'E':
        this.selectedCategory = 'Expense';
        break;
      case 'A':
        this.selectedCategory = 'Asset';
        break;
      case 'L':
        this.selectedCategory = 'Liability';
        break;
      case 'Q':
        this.selectedCategory = 'Equity';
        break;
    }

    // Refresh data using your original API
    this.getGLAccountsInfo();
  }
  toggleSelectAll() {
    if (this.storeIds.length === this.stores.length) {
      this.storeIds = [];
    } else {
      this.storeIds = this.stores.map((s) => s.ID);
    }
    this.updateStoreSelection();
  }
  onStoreToggle(storeId: number) {
    const idx = this.storeIds.indexOf(storeId);
    if (idx >= 0) this.storeIds.splice(idx, 1);
    else this.storeIds.push(storeId);

    this.updateStoreSelection();
  }
  selectAllStores(): void {
    this.storeIds = this.stores.map((s) => s.ID);
    this.updateStoreButtonLabel();
  }
  updateStoreButtonLabel(): void {
    if (this.storeIds.length === this.stores.length) {
      this.storeButtonLabel = 'All';
    } else if (this.storeIds.length === 1) {
      const selected = this.stores.find((s) => s.ID === this.storeIds[0]);
      this.storeButtonLabel = selected?.storename || '';
    } else {
      this.storeButtonLabel = this.storeIds.length.toString();
    }
  }
  allstores(state: 'Y' | 'N') {
    this.storeIds = state === 'Y' ? this.stores.map((s) => s.ID) : [];
    this.updateStoreSelection();
  }
  updateStoreSelection() {
    if (this.storeIds.length === this.stores.length) {
      this.storeButtonLabel = 'All';
      this.selectedVars.var1 = 'All';
    } else if (this.storeIds.length === 1) {
      const selected = this.stores.find((s) => s.ID === this.storeIds[0]);
      this.storeButtonLabel = selected?.storename || '';
      this.selectedVars.var1 = selected?.storename || '';
    } else {
      this.storeButtonLabel = this.storeIds.length.toString();
      this.selectedVars.var1 = this.storeIds.join(',');
    }
  }
  individualStores(store: any) {
    this.onStoreToggle(store.ID);
  }

  frommonth: Date | null = null;
  Dupfrommonth: Date | null = null;

  minFromMonth: Date = new Date(2000, 0, 1);
  maxFromMonth: Date = new Date();

  monthPickerConfig: Partial<BsDatepickerConfig> = {
    dateInputFormat: 'MMMM YYYY',
    minMode: 'month' as 'month',
    showWeekNumbers: false,
    minDate: this.minFromMonth,
    maxDate: this.maxFromMonth,
  };
  changeDatefrom(e: Date | null) {
    this.activePopover = -1;
    if (!e) return;

    this.frommonth = e;

    const year = e.getFullYear();
    const month = e.getMonth();
    this.fromDate = new Date(year, month, 1);
    const lastDay = new Date(year, month + 1, 0);
    this.toprevious = e;
    this.minToMonth = e;
    if (this.tomonth && this.tomonth < e) {
      this.tomonth = lastDay;
    }
  }

  // filterMapped(mapped: boolean) {
  //   this.isMapped = mapped;
  //   this.applyFilters();
  // }

  getGLAccountsInfo(accountNumber: string = '') {
    this.shared.spinner.show();
    this.AccountingMapping = [];
    this.NoData = false;
    const body = {
      map_type: this.statusFilter[0],
      store_Id: this.storeIds.length ? this.storeIds.join(',') : '0',
      account_Type: this.mapCategoryToType(this.selectedCategory),
      accountnumber: accountNumber,
      UserID: 1,
      Date: this.frommonth ? this.formatDateForApi(this.frommonth) : '10-12-2025',
    };

    console.log('API Request Body:', body);

    this.shared.api
      .postmethod(this.shared.common.routeEndpoint + 'GetGLAccountsInfo', body)
      .subscribe({
        next: (response: any) => {
          if (response?.status === 200) {
            const rows = response?.response || [];

            // If API returned rows
            if (rows.length > 0) {
              console.log('API Response:', rows);

              this.tableData = rows; // raw data
              this.filteredTableData = [...rows]; // filtered copy
              this.AccountingMapping = [...rows]; // grid binding
              this.shared.spinner.hide();

              this.NoData = false;
            } else {
              // Empty response
              this.NoData = true;
              this.tableData = [];
              this.filteredTableData = [];
              this.AccountingMapping = [];
              this.shared.spinner.hide();
            }
          } else {
            // Invalid response
            this.NoData = true;
            this.tableData = [];
            this.filteredTableData = [];
            this.AccountingMapping = [];
          }

          // Pagination logic after filters
          this.Pagination = this.AccountingMapping.length > 10;

          this.cdr.detectChanges();
          //  this.shared.spinner.hide();
        },

        error: (err: any) => {
          console.error('API Error:', err);
          this.shared.spinner.hide();

          this.NoData = true;
          this.tableData = [];
          this.filteredTableData = [];
          this.AccountingMapping = [];
        },
      });
  }

  searchFilter() {
    this.searchTerm = this.searchText.trim(); // sync with backend usage
    this.getGLAccountsInfo(this.searchTerm);
  }

  onSearchInput() {
  // if search field becomes empty, reload original API
  if (!this.searchText.trim()) {
    this.searchTerm = '';
    this.getGLAccountsInfo('');
  }
}

  StatusFiltering(type: string) {
    this.statusFilter = [type];

    this.getGLAccountsInfo(this.searchTerm);
  }

  applyFilters() {
    this.activePopover = -1;
    this.isPopoverOpen = false;

    let data = this.tableData;

    // apply date filter if needed
    if (this.frommonth) {
      const year = this.frommonth.getFullYear();
      const month = this.frommonth.getMonth();
      this.fromDate = new Date(year, month, 1);
    }

    if (this.tomonth) {
      const year = this.tomonth.getFullYear();
      const month = this.tomonth.getMonth();
      this.toDate = new Date(year, month + 1, 0);
    }

    this.filteredTableData = data;
    // update AccountingMapping used in template
    this.AccountingMapping = [...this.filteredTableData];
    this.Pagination = this.AccountingMapping.length > 10;
    this.NoData = this.AccountingMapping.length === 0;

    this.cdr.detectChanges();
  }

  formatDateForApi(date: Date): string {
    const d = new Date(date);
    const month = ('0' + (d.getMonth() + 1)).slice(-2);
    const day = ('0' + d.getDate()).slice(-2);
    const year = d.getFullYear();
    return `${month}-${day}-${year}`;
  }

  mapCategoryToType(category: string): string {
    switch (category) {
      case 'Sale':
        return 'S';
      case 'Cost':
        return 'C';
      case 'Non-Operating Expense':
        return 'X';
      case 'Non-Operating Income':
        return 'I';
      case 'Expense':
        return 'E';
      case 'Asset':
        return 'A';
      case 'Liability':
        return 'L';
      case 'Equity':
        return 'Q';
      default:
        return 'S';
    }
  }

  setCategory(category: string) {
    this.selectedCategory = category;
    this.getGLAccountsInfo();
  }

  // -------------- CLIENT A MAPPING FUNCTIONS -----------------

  /**
   * bulkEdit(BulkFlag, mode, p3, p4, am, actionType)
   * This mirrors Client A function signature and behavior.
   * - BulkFlag: whether multiple records are being edited
   * - am: record or array of records
   */
  bulkEdit(e: any, ref: any, template: any, multiorsingle: any, record: any, BI?: any) {
    this.BIref = BI;

    if (BI === 'MappedEditMapping') {
      this.Refchange = 'editmapping';
      this.bulkEditRecords = [];
      this.submitted = false;

      // Identify Single or Multiple Selection
      if (ref === 'E') {
        this.Bulk = true;
        this.ignorebtn = true;
      } else {
        if (multiorsingle === 'S') {
          // Insert selected record
          this.bulkEditRecords.push(record);

          // Load all fields + dropdown values just like Client A
          this.OnChange('Default', '', '', '', '', '', '', '', '', '', '');

          // Reset all local dropdown selections
          if (this.bulkEditRecords.length > 0) {
            this.accounttype = '0';
            this.selectedaccounttypedetail = '0';
            this.departmenttype = '0';
            this.department = '0';
            this.subtype = '0';
            this.subtypedetail = '0';
            this.subtypefulldetail = '0';
            this.flexcol1 = '0';
            this.financialsummary = '0';

            // Empty all dependent dropdown lists
            this.Accounttypedetail = [];
            this.Departmenttype = [];
            this.Department = [];
            this.Subtype = [];
            this.SubtypeDetail = [];
            this.SubtypeFullDetail = [];
            this.Flexcol1 = [];
            this.Financialsummary = [];

            // Flags reset just like Client A
            // this.Asubtype = false;
            // this.Asubtypedetail = false;
            // this.Asubtypefulldetail = false;
          }
        }
      }
    }
  }

  OnChange(
    ref: string,
    storeId?: any,
    accountNo?: any,
    accountTypeRaw?: any,
    accountTypeDetail?: any,
    opArea?: any,
    dept?: any,
    subtype?: any,
    subtypeDetail?: any,
    subtypeFullDetail?: any,
    flexcol1?: any,
    refChange?: any,
    actbal?: any
  ) {
    if (refChange === 'AE') {
      this.Refchange = '';
      this.resetAllSelections();
    }

    const body = {
      FilterType: ref === 'AT' ? this.actbal : ref,
      store_ID: this.bulkEditRecords[0].companyId,
      accountnumber: this.bulkEditRecords[0].accountNumber,
      ACCOUNTTYPERAW: accountTypeRaw,
      ACCOUNTTYPEDETAIL: accountTypeDetail,
      OperationalArea: opArea,
      DEPARTMENT: dept,
      SUBTYPE: subtype,
      SUBTYPEDETAIL: subtypeDetail,
      SUBTYPEFULLDETAIL: subtypeFullDetail,
      FLEXCOL1: flexcol1,
    };

    console.log(accountTypeRaw,'Account Type Raw');
    this.shared.api
      .postmethod(this.shared.common.routeEndpoint + 'GetAccountMappingFiltersV2', body)
      .subscribe((res: any) => {
        if (res.status !== 200) return;

        const list = res.response;

        /* -----------------------
         DEFAULT (initial load)
      ------------------------*/
        if (ref === 'Default') {
          this.resetAllSelections();
          this.defaultresponse = res;

          const first = list.find((e: any) => e.OperationalType);
          this.actbal = first?.OperationalType || 'Activity';

          this.OnChange('AT');
          return;
        }

        /* -----------------------
         AT — Account Type
      ------------------------*/
        if (ref === 'AT') {
          this.resetLowerLevels('AT');
          this.Accounttype = list.filter((e: { accountType: any }) => e.accountType);

          if (this.Refchange === 'editmapping') {
            this.accounttype = this.bulkEditRecords[0].accountType;
          } else {
            // auto-select if only one
            if (this.Accounttype.length === 1) {
              this.accounttype = this.Accounttype[0].accountType;
            } else {
              this.accounttype = '0';
            }
          }

          // trigger next level
          if (this.actbal === 'Activity') this.OnChange('OA', '', '', this.accounttype);
          else this.OnChange('ATD', '', '', this.accounttype);

          return;
        }

  //        // ALWAYS pass accountType
  // if (this.actbal === 'Activity'){
  //   this.OnChange('OA', '', '', this.accounttype);
  // } else {
  //   this.OnChange('ATD', '', '', this.accounttype);
  //    return;
  // }
 


        /* -----------------------
         ATD — Account Type Detail
      ------------------------*/
        if (ref === 'ATD') {
          this.resetLowerLevels('ATD');
          this.Accounttypedetail = list.filter(
            (e: { AccountTypeDetail: any }) => e.AccountTypeDetail
          );

          if (this.Refchange === 'editmapping') {
            this.selectedaccounttypedetail = this.bulkEditRecords[0].acctSubtype;
          } else {
            this.selectedaccounttypedetail =
              this.Accounttypedetail.length === 1
                ? this.Accounttypedetail[0].AccountTypeDetail
                : '0';
          }

          this.OnChange('OA', '', '', this.accounttype);
          return;
        }

        /* -----------------------
         OA — Operational Area
      ------------------------*/
        if (ref === 'OA') {
          console.log(this.accounttype,'Account Type');
          
          this.resetLowerLevels('OA');
          this.Departmenttype = list.filter((e: { OperationalArea: any }) => e.OperationalArea);

          if (this.Refchange === 'editmapping') {
            this.departmenttype = this.bulkEditRecords[0].operationalArea;
          } else {
            this.departmenttype =
              this.Departmenttype.length === 1 ? this.Departmenttype[0].OperationalArea : '0';
          }

          this.OnChange('DT', '', '', this.accounttype, '', this.departmenttype);
          return;
        }

        /* -----------------------
         DT — Department
      ------------------------*/
        if (ref === 'DT') {
          this.resetLowerLevels('DT');
          this.Department = list.filter((e: any) => e.Department);

          if (this.Refchange === 'editmapping') {
            this.department = this.bulkEditRecords[0].department;
          } else {
            this.department = this.Department.length === 1 ? this.Department[0].Department : '0';
          }

          this.OnChange('ST', '', '', this.accounttype, '', '', this.department);

          return;
        }

        /* -----------------------
         ST — SubType
      ------------------------*/
        if (ref === 'ST') {
          this.resetLowerLevels('ST');
          this.Subtype = list.filter((e: { SubType: any }) => e.SubType);

          if (this.Refchange === 'editmapping') {
            this.subtype = this.bulkEditRecords[0].subType;
          } else {
            this.subtype = this.Subtype.length === 1 ? this.Subtype[0].SubType : '0';
          }

          this.finsummary();
          this.OnChange('STD', '', '', this.accounttype, '', '', this.department, this.subtype);
          return;
        }

        /* -----------------------
         STD — SubType Detail
      ------------------------*/
        if (ref === 'STD') {
          this.resetLowerLevels('STD');
          this.SubtypeDetail = list.filter((e: { SubTypeDetail: any }) => e.SubTypeDetail);

          if (this.Refchange === 'editmapping') {
            this.subtypedetail = this.bulkEditRecords[0].subTypeDetail;
          } else {
            this.subtypedetail =
              this.SubtypeDetail.length === 1 ? this.SubtypeDetail[0].SubTypeDetail : '0';
          }

          this.OnChange(
            'STFD',
            '',
            '',
            this.accounttype,
            '',
            '',
            this.department,
            this.subtype,
            this.subtypedetail
          );
          return;
        }

        /* -----------------------
         STFD — SubType Full Detail
      ------------------------*/
        if (ref === 'STFD') {
          this.SubtypeFullDetail = list.filter((e: any) => e.SubTypeFullDetail);
          return;
        }

        /* -----------------------
         FS — Financial Summary
      ------------------------*/
        if (ref === 'FS') {
          this.Financialsummary = list.filter((e: any) => e.Fin_Summary);
          
        }
         else {
  this.financialsummary =
    this.Financialsummary.length === 1
      ? this.Financialsummary[0].Fin_Summary_Key
      : '0';
}

        /* -----------------------
         BS — Balance Summary
      ------------------------*/
        if (ref === 'BS') {
          this.Balancesummary = list.filter((e: any) => e.Balance_Subtype);
          return;
        }
      });
  }

  finsummary() {
    const obj = {
      Accounttype: this.accounttype,
      Department: this.department,
    };
    // this.apiSrvc
    //   .postmethod

    this.shared.api
      .postmethod(this.shared.common.routeEndpoint + 'GetAccountFinSummary', obj)
      .subscribe((res: any) => {
        if (!res) return;
        // guard and normalize response before accessing properties
        const resp: any = res;
        if (resp.status == 200) {
          const responseArray = resp.response || [];
          this.Financialsummary = responseArray.filter(
            (e: any) =>
              e.Fin_Summary_Key != '' && e.Fin_Summary_Key != null && e.Fin_Summary_Key != 0
          );

          console.log(this.Financialsummary);
          if (this.Financialsummary.length == 1) {
            if (this.Refchange == 'editmapping') {
              this.financialsummary = this.bulkEditRecords[0].Fin_Summary;
            } else {
              this.financialsummary = this.Financialsummary[0].Fin_summary;
            }
          } else {
            if (this.Refchange == 'editmapping') {
              this.financialsummary = this.bulkEditRecords[0].Fin_Summary;
            } else {
              this.financialsummary = '0';
            }
          }
        }
      });
  }

  resetAllSelections() {
    this.accounttype = '0';
    this.selectedaccounttypedetail = '0';
    this.departmenttype = '0';
    this.department = '0';
    this.subtype = '0';
    this.subtypedetail = '0';
    this.subtypefulldetail = '0';
    this.flexcol1 = '0';
  }

  resetLowerLevels(level: string) {
    if (level === 'AT') {
      this.selectedaccounttypedetail = '0';
      this.departmenttype = '0';
      this.department = '0';
      this.subtype = '0';
      this.subtypedetail = '0';
      this.subtypefulldetail = '0';
      this.flexcol1 = '0';
      return;
    }

    if (level === 'ATD') {
      this.departmenttype = '0';
      this.department = '0';
      this.subtype = '0';
      this.subtypedetail = '0';
      this.subtypefulldetail = '0';
      this.flexcol1 = '0';
      return;
    }

    if (level === 'OA') {
      this.department = '0';
      this.subtype = '0';
      this.subtypedetail = '0';
      this.subtypefulldetail = '0';
      this.flexcol1 = '0';
      return;
    }

    if (level === 'DT') {
      this.subtype = '0';
      this.subtypedetail = '0';
      this.subtypefulldetail = '0';
      this.flexcol1 = '0';
      return;
    }

    if (level === 'ST') {
      this.subtypedetail = '0';
      this.subtypefulldetail = '0';
      this.flexcol1 = '0';
      return;
    }

    if (level === 'STD') {
      this.subtypefulldetail = '0';
      this.flexcol1 = '0';
      return;
    }

    if (level === 'STFD') {
      this.flexcol1 = '0';
      return;
    }
  }

 closeBootstrapModal() {
  const modalEl = document.getElementById('mapModal');
  if (!modalEl) return;

  const modal = bootstrap.Modal.getInstance(modalEl) 
              || new bootstrap.Modal(modalEl);  

  modal.hide();

  // clean backdrop & body lock
  document.querySelectorAll('.modal-backdrop').forEach(el => el.remove());
  document.body.classList.remove('modal-open');
  document.body.style.overflow = '';
}

  save(): void {
  this.submitted = true;

  this.missingFields = this.getMissingFields();
  if (this.missingFields.length > 0) {
    this.modalService.open(this.Alert, { centered: true });
    return;
  }

  const payload = {
    companyid: this.bulkEditRecords[0].companyId,
    accountnumber: this.bulkEditRecords[0].accountNumber,
    accountdescription: this.bulkEditRecords[0].accountDescription,
    accounttype: this.accounttype,
    accountTypeDetail: this.actbal == 'Activity' ? '' : this.selectedaccounttypedetail,
    operationalArea: this.departmenttype,
    department: this.department,
    subType: this.subtype,
    subTypeDetail: this.subtypedetail,
    subTypeFullDetail: '',
    OperationalType: this.actbal,
    flexCol1: '',
    flexCol2: '',
    finSummary: this.actbal == 'Balance' ? '' : this.financialsummary,
  };

  this.shared.spinner.show();

  this.shared.api.postmethod('accountmapping/AccountMappingActionV2', payload).subscribe({
    next: (res: any) => {
      this.shared.spinner.hide();

      if (res && (res.status ==200 )) {
        alert('This Account Mapped successfully');

        this.getGLAccountsInfo();
//  this.ngbmodalActive.dismiss(); // close modal
 (<HTMLInputElement>document.getElementById('forceCloseModal')).click();
 
        this.onclose();
      } else {
        alert(res?.message || 'Failed to save mapping');
      }
    },
    error: (err) => {
      this.shared.spinner.hide();
      console.error('UpdateAccountMapping error', err);
      alert('Error saving mapping');
    },
  });
}

  MapAccountCancel(): void {
    this.submitted = false;
    this.onclose();
  }

  onclose(): void {
    // Reset modal fields and state
    this.bulkEditRecords = [];
    this.actbal = '';
    this.accounttype = 0;
    this.selectedaccounttypedetail = 0;
    this.departmenttype = 0;
    this.department = 0;
    this.subtype = 0;
    this.subtypedetail = 0;
    this.financialsummary = 0;
    this.submitted = false;

    // programmatically dismiss modal if using NgbModal
    try {
      this.modalService.dismissAll();
    } catch (err) {
      // ignore if not using NgbModal programmatic control
    }
  }

  // ---------- Pagination helpers ----------
  Prev(): void {
    if (this.PageCount > 1) {
      this.PageCount--;
      this.getGLAccountsInfo();
    }
  }

  Next(): void {
    if (this.LastCount) {
      this.PageCount++;
      this.getGLAccountsInfo();
    }
  }

  getMissingFields(): string[] {
    const fields = [];

    if (!this.actbal || this.actbal === '0') fields.push('Account Balance');
    if (!this.accounttype || this.accounttype === '0') fields.push('Account Type');
    if (!this.departmenttype || this.departmenttype === '0') fields.push('Operational Area');
    if (!this.department || this.department === '0') fields.push('Department');
    if (!this.subtype || this.subtype === '0') fields.push('Sub Type');
    if (!this.subtypedetail || this.subtypedetail === '0') fields.push('Sub Type Detail');

    return fields;
  }

  ngAfterViewInit(): void {
    this.excel = this.shared.api.getExportToExcelAllReports().subscribe((res: any) => {
      if (res && res.obj && res.obj.title == 'Account Mapping' && res.obj.state == true) {
        this.exportToExcel(); // merged export will create both sheets
      }
    });
     const modalEl = document.getElementById('mapModal');
  if (modalEl) {
    this.modalInstance = new bootstrap.Modal(modalEl);
  }
  }
  exportToExcel() {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet('Account Mapping');

    worksheet.addRow([]);

    // Title
    const titleRow = worksheet.addRow(['ACCOUNT MAPPING REPORT']);
    worksheet.mergeCells('A2:L2');
    titleRow.font = { bold: true, size: 15 };
    titleRow.alignment = { horizontal: 'center' };

    worksheet.addRow([]);
    // ========================= FILTER SECTION =========================

    // REPORT MONTH
    const reportMonth = this.frommonth
      ? this.frommonth.toLocaleString('en-US', { month: 'long', year: 'numeric' })
      : '-';

    worksheet.addRow(['Report Month:', reportMonth]);
    if (worksheet.lastRow) worksheet.lastRow.getCell(1).font = { bold: true };

    // CATEGORY
    worksheet.addRow(['Category:', this.selectedCategory || '-']);
    if (worksheet.lastRow) worksheet.lastRow.getCell(1).font = { bold: true };

    // MAPPED / UNMAPPED
    const mapStatus = this.statusFilter[0] === 'M' ? 'Mapped' : 'Unmapped';
    worksheet.addRow(['Mapped Type:', mapStatus]);
    if (worksheet.lastRow) worksheet.lastRow.getCell(1).font = { bold: true };

    // SEARCH TERM
    worksheet.addRow(['Search:', this.searchText || '-']);
    if (worksheet.lastRow) worksheet.lastRow.getCell(1).font = { bold: true };
    // ---------------------- STORE NAME SECTION ----------------------
    // let storeNames: any[] = [];
    // let store = this.storeIds;

    // const selectedStores = this.stores.filter((s: any) =>
    //   store.includes(s.ID)
    // );

    // if (store.length === this.stores.length) {
    //   storeNames = ["All Stores"];
    // } else {
    //   storeNames = selectedStores.map((s: any) => s.storename);
    // }

    // const storeRow = worksheet.addRow(["Stores:", storeNames.join(", ")]);
    // storeRow.font = { bold: true };
    // storeRow.getCell(2).font = { color: { argb: "0066CC" } };

    // worksheet.addRow([]);

    // ---------------------- TABLE HEADERS ----------------------
    const headers = [
      'Store',
      'Acct Desc',
      'Acct Number',
      'Account Type',
      // 'Account Type Detail',
      'Operational Area',
      'Department',
      'Sub Type',
      'Sub Type Detail',
      'Fin Summary',
      'MTD',
      'YTD',
    ];

    const headerRow = worksheet.addRow(headers);

    headerRow.font = { bold: true, color: { argb: 'FFFFFF' }, size: 10 };
    headerRow.alignment = { horizontal: 'center', vertical: 'middle' };

    headerRow.eachCell((cell) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '2a91f0' },
      };
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      };
    });

    worksheet.views = [{ state: 'frozen', ySplit: 4 }];

    // ---------------------- COLUMN WIDTHS ----------------------
    worksheet.columns = [
      { width: 25 },
      { width: 35 },
      { width: 15 },
      { width: 15 },
      // { width: 20 },
      { width: 20 },
      { width: 20 },
      { width: 20 },
      { width: 25 },
      { width: 20 },
      { width: 15 },
      { width: 15 },
    ];

    // ---------------------- DATA ROWS ----------------------
    let rowIndex = 0;

    this.AccountingMapping.forEach((item: any) => {
      const row = worksheet.addRow([
        item.as_companyName || '-',
        item.accountDescription || '-',
        item.accountNumber || '-',
        item.accountType || '-',
        // item.acctSubtype || '-',
        item.operationalArea || '-',
        item.department || '-',
        item.subType || '-',
        item.subTypeDetail || '-',
        item.Fin_Summary || '-',
        item.MTD === null ? '-' : item.MTD,
        item.YTD === null ? '-' : item.YTD,
      ]);

      row.eachCell((cell) => {
        cell.border = {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' },
        };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      });

      // Currency formatting
      row.getCell(11).numFmt = '$#,##0.00';
      row.getCell(12).numFmt = '$#,##0.00';

      // Zebra Striping
      if (rowIndex % 2 === 1) {
        row.eachCell((cell) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'e5e5e5' },
          };
        });
      }

      rowIndex++;
    });

    // ---------------------- DOWNLOAD FILE ----------------------
    workbook.xlsx.writeBuffer().then((buffer) => {
      const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });

      FileSaver.saveAs(blob, 'Account_Mapping_Report.xlsx');
    });
  }
}
