

import { Component, ElementRef, ViewChild, HostListener, OnInit, EventEmitter, Output, Input, Renderer2 } from '@angular/core';
import { NgbActiveModal, NgbDateParserFormatter, NgbModal } from '@ng-bootstrap/ng-bootstrap';
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
import { Stores } from '../../../../CommonFilters/stores/stores';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
import FileSaver from 'file-saver';
// import { ToastrService } from 'ngx-toastr';

// declare let require: any;
// const pdfMake = require('pdfmake/build/pdfmake');
// const pdfFonts = require('pdfmake/build/vfs_fonts');
// (pdfMake as any).vfs =
//   (pdfFonts as any).vfs ||
//   ((pdfFonts as any).pdfMake && (pdfFonts as any).pdfMake.vfs);

const EXCEL_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
const EXCEL_EXTENSION = '.xlsx';

@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [SharedModule, BsDatepickerModule, TimeConversionPipe, Stores],
  templateUrl: './dashboard.html',
  styleUrls: ['./dashboard.scss'],
})
export class Dashboard implements OnInit {
  // Dates & filters (appointment-style)
  FromDate: any;
  ToDate: any;
  CurrentDate = new Date();

  // user / store
  storeIds: any = '';
  Role: any = [];
  userid: any;

  // UI state
  spinnerLoader: boolean = true;
  enablevehicle: any = false;
  vehiclear: any = 'WOAR';
  noData: boolean = false;
  NoData: boolean = false;
  days: any = [];
  LogCount = 1;
  solutionurl: any = (window as any)['environment']?.apiUrl || '';
  groups: any = 1;
  callLoadingState = 'FL';
  check: boolean = false;
  Viewmore: boolean = false;

  // Data
  FloorPlanData: any = [];
  FloorPlanTotalData: any = [];
  TotalFloorPlanData: any = [];
  QISearchName: any = '';
  commentdata: any = [];
  notesstage: any = [];

  // filter defaults
  dealStatus: any = ['Booked', 'Finalized', 'Delivered'];
  saleType: any = ['Retail', 'Fleet'];
  tempPriorityFilter: any = ['All'];
  allordebit: any = 'dr';
  financeManagerId: any = '0';
  priorityRecords: any = [];
  priorityVisibility: boolean = false;
  prioritycheck: boolean = false;
  originalFloorPlanData: any[] = [];
  // header & widgets
  header: any = [
    {
      type: 'Bar',
      storeIds: this.storeIds,
      vehiclear: this.vehiclear,
      dealStatus: this.dealStatus,
      saleType: this.saleType,
      allordebit: this.allordebit,
      groups: this.groups,
      financemanagers: this.financeManagerId,
    },
  ];
  popup: any = [{ type: 'Popup' }];

  // notes / hide records
  notesStageValue: any = '';
  notesStageValueGrid: any = '';
  notesStageText: any = '';
  selecteddata: any = [];
  FinalArray: any = [];
  hideRecords: any = [];

  // misc
  today: any;
  startDate: any;
  endDate: any;
  column: string = 'CategoryName';
  isDesc: boolean = false;
  callExportState!: Subscription;
  reportOpenSub!: Subscription;
  reportGetting!: Subscription;
  Pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;
  CITstate: any;

  // view helpers
  public isCollapsed: boolean = false;
  public isCollapsable: boolean = false;
  maxHeight: number = 10;
  Favreports: any = [];

  // notes view / comments
  notesViewState: boolean = true;
  commentsVisibility: boolean = false;
  hideVisibility: boolean = false;

  // scroll
  @ViewChild('scrollcent') scrollcent!: ElementRef;
  Scrollpercent: any = 0;
  scrollCurrentposition: any = 0;

  StoreVal: any = '71,53,8,7,4,35,1,32,40,50,25,18,31,3,70,72,2,17,41,55,42,51,12,73,54,9,15,5,14,30,11';
  StoresIds: any = [];

  Performance: any = 'Load';
  maxDate: any;
  // vehiclear already defined above (string); keep an array for the report controls too:
  vehiclearArray: any = ['WAR'];
  @Input() headerInput: any;
  @Input() popupInput: any;
  Bar: boolean = false;
  storeName: any = '';
  employeeschanges: any = '';
  @Input() QISearchNameInput: any;
  @Output() messageEvent = new EventEmitter<string>();
  selectedstorevalues: any = [];
  AllStores: boolean = true;
  selectedGroups: any = [];
  AllGroups: boolean = true;
  groupstate: boolean = false;

  month: any;
  selectedFiManagersvalues: any = [];
  selectedFiManagersname: any = [];
  financeManager: any = [];
  helpdata: any;
  activePopover: number = -1;
  bsConfig!: Partial<BsDatepickerConfig>;

  groupsArray: any = [];
  storename: any = ''
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  groupId: any = 0;
  stores: any = []
  //  storeIds: any = '';
  storesFilterData: any = {
    'groupsArray': this.groupsArray, 'groupId': this.groupId, 'storesArray': this.stores, 'storeids': '1', 'type': 'M', 'others': 'N',
    'groupName': this.groupName, 'storename': this.storename, storecount: null, 'storedisplayname': this.storedisplayname
  };
  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent) {
    const clickedInside = (event.target as HTMLElement).closest('.dropdown-toggle, .dropdown-menu , .timeframe, .reportstores-card');
    if (!clickedInside) {
      this.activePopover = -1;
    }
  }

  constructor(
    public shared: Sharedservice,
    public setdates: Setdates,
    private ngbmodal: NgbModal,
    private ngbmodalActive: NgbActiveModal,
    private elementRef: ElementRef,
    private renderer: Renderer2,
    private cdRef: ChangeDetectorRef,
    private datepipe: DatePipe,
    public formatter: NgbDateParserFormatter,
    public toast: ToastService,

  ) {
    // set defaults and mirrored logic from appointment style
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
    localStorage.setItem('time', 'C');

    this.commentsVisibility = true;

    // init user/store info from local storage (same logic)
    if (localStorage.getItem('userInfo') != null) {
      // Keep same logic but don't break when redirectionFrom missing
      try {
        const ud: any = JSON.parse(localStorage.getItem('userInfo')!);
        this.Role = ud.user_Info.title;
        this.userid = ud.user_Info.userid;
        console.log(this.Role, this.userid);
      } catch {
        // ignore
      }
      // if comm.redirectionFrom exists apply original logic
      // try {
      const ud: any = JSON.parse(localStorage.getItem('userInfo')!);
      let allordebit = ud.flag2
      console.log(ud, allordebit, '...................................................................')
      if (ud.user_Info.flag != 'M') {
        this.dealStatus = this.dealStatus;
        this.financeManagerId = 0;
        this.allordebit = allordebit == 'A' ? 'all' : 'dr';
      }
      // } catch { /* ignore */ }
    }

    // date range for UI (same logic)
    this.today = new Date();
    this.startDate = new Date();
    this.endDate = new Date(this.startDate);
    this.endDate = new Date(this.endDate.setDate(this.endDate.getDate() + 4));
    for (let q = new Date(this.startDate); q <= this.endDate; q.setDate(q.getDate() + 1)) {
      this.days.push({ day: q.toString() });
    }



    // if (this.shared.common.groupsandstores.length > 0) {
    //   this.groupsArray = this.shared.common.groupsandstores.filter((val: any) => val.sg_id != this.shared.common.reconID);
    //   this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
    //   this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_Name : this.groupName = ''
    //   this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
    //   this.getStoresandGroupsValues()
    // }
    // set page title
    try {
      this.shared.setTitle(this.shared.common.titleName + '-CIT');
    } catch {
      // fallback: no-op
    }

    // initialize From/To based on original logic
    this.CurrentDate.setDate(this.CurrentDate.getDate() - 1);
    this.ToDate = new Date(
      this.CurrentDate.getFullYear(),
      this.CurrentDate.getMonth(),
      2
    );
    this.FromDate = this.ToDate.toISOString().slice(0, 10);
    this.ToDate = this.CurrentDate.toISOString().slice(0, 10);

    // set header data unless favorite
    if (localStorage.getItem('Fav') != 'Y') {
      const data = {
        title: 'CIT',
        stores: this.storeIds,
        groups: this.groups,
        financemanagers: this.financeManagerId,
        dealStatus: this.dealStatus,
        saleType: this.saleType,
        allordebit: this.allordebit,
        vehiclear: this.vehiclear,
      };
      try {
        this.shared.api.SetHeaderData({ obj: data });
      } catch { /* ignore */ }

      this.header = [
        {
          type: 'Bar',
          storeIds: this.storeIds,
          vehiclear: this.vehiclear,
          dealStatus: this.dealStatus,
          saleType: this.saleType,
          allordebit: this.allordebit,
          financemanagers: this.financeManagerId,
          groups: this.groups,
        },
      ];

      if (this.storeIds != '') {
        this.Getfloorplansdata();
        this.getEmployees();
      }
    }

    // listen window clicks for certain modal close behavior (from report component)
    this.renderer.listen('window', 'click', (e: Event) => {
      const TagName = e.target as HTMLButtonElement;
      if (TagName && TagName.className === 'd-block modal fade show modal-static') {
        try { this.ngbmodal.dismissAll(); } catch { }
      }
    });
  }

  ngOnInit(): void {
    // datepicker config and other init tasks
    this.bsConfig = {
      dateInputFormat: 'MM/dd/yyyy',
      showWeekNumbers: false,
      adaptivePosition: true
    };
  }

  // Messaging (report -> parent)
  sendMessage() {
    this.messageEvent.emit(this.QISearchName);
  }

  /* ------------------------------
     Core: Get floor plan / CIT data
     ------------------------------ */
  selectedAgingFrom: number | null = 0;
  selectedAgingTo: number | null = 0;
  agingFilter: { min: number; max: number | null } | null = null;
  applyAgingFilter(min: number, max: number | null) {
    this.agingFilter = { min, max };
    this.selectedAgingFrom = min;
    this.selectedAgingTo = max;

    console.log(this.agingFilter)
    this.Getfloorplansdata()

  }
  Getfloorplansdata() {
    this.NoData = false;
    this.FloorPlanData = [];
    this.FloorPlanTotalData = [];
    this.filteredFloorplanData = [];
    try { this.shared.spinner.show(); } catch { /* ignore */ }

    const obj = {
      AS_ID: this.storeIds.toString(),
      DealStatus: this.dealStatus.toString(),
      ValueType: this.allordebit && this.allordebit.toString ? (this.allordebit.toString() == 'all' ? '' : this.allordebit.toString()) : '',
      FIManagerID: this.financeManagerId,
      UserID: 0,
      GetPriority: this.tempPriorityFilter && this.tempPriorityFilter.toString ? (this.tempPriorityFilter.toString() == 'All' ? '' : this.tempPriorityFilter.toString()) : '',
      // Aging_From: this.agingFilter?.min ?? null,
      // Aging_To: this.agingFilter?.max ?? null,
      // Retail_Fleet: this.saleType.toString(),
    };

    this.shared.api.postmethod(this.shared.common.routeEndpoint + 'GetCITFloorPlanData', obj).subscribe(
      (res: any) => {
        try { this.shared.spinner.hide(); } catch { }
        if (res && res.status == 200) {
          if (res.response && res.response.length > 0) {
            this.FloorPlanData = res.response.filter((e: any) => e.store != 'TOTAL');
            this.FloorPlanTotalData = res.response.filter((e: any) => e.store == 'TOTAL');

            // Parse AgeData / Comments / Notes safely
            this.FloorPlanData.forEach((x: any) => {
              if (x && x.AgeData && typeof x.AgeData === 'string') {
                try { x.AgeData = JSON.parse(x.AgeData); } catch { }
              }
              if (x && x.Comments && typeof x.Comments === 'string') {
                try { x.Comments = JSON.parse(x.Comments); } catch { }
              }
              if (x && x.Notes && typeof x.Notes === 'string') {
                try {
                  x.Notes = JSON.parse(x.Notes);
                  x.duplicateNotes = x.Notes;
                  x.Notesstate = '+';
                  if (x.Notes.length > 3) {
                    x.duplicateNotes = x.duplicateNotes.slice(0, 3);
                  } else {
                    x.duplicateNotes = x.Notes;
                  }
                } catch { }
              }
            });
            this.originalFloorPlanData = JSON.parse(
              JSON.stringify(this.FloorPlanData)
            );
            this.filteredFloorplanData = this.FloorPlanData || [];
            this.NoData = this.FloorPlanData.length === 0;
          } else {
            this.NoData = true;
          }
        } else {
          this.NoData = true;
        }
      },
      (error: any) => {
        try { this.shared.spinner.hide(); } catch { }
        this.NoData = true;
      }
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
        x.AgeDate?.toString().toLowerCase().includes(term) ||
        x.Account?.toString().toLowerCase().includes(term) ||
        x.Control?.toString().toLowerCase().includes(term) ||
        x.Control2?.toString().toLowerCase().includes(term) ||

        x.BalCIT?.toString().toLowerCase().includes(term) ||
        x.BalFloorplan?.toString().toLowerCase().includes(term) ||

        x.CustomerName?.toLowerCase().includes(term) ||
        x.CustomerNumber?.toString().toLowerCase().includes(term) ||

        x.store?.toLowerCase().includes(term) ||

        x.StockNumner?.toString().toLowerCase().includes(term) ||
        x.Deal?.toString().toLowerCase().includes(term) ||

        x.Stage?.toLowerCase().includes(term) ||
        x.FIManager?.toLowerCase().includes(term) ||
        x.SalesManager?.toLowerCase().includes(term) ||
        x.SP1?.toLowerCase().includes(term) ||

        x.SaleType?.toLowerCase().includes(term) ||
        x.DealType?.toLowerCase().includes(term) ||
        x.DealStatus?.toLowerCase().includes(term) ||

        x.BankName?.toLowerCase().includes(term) ||

        x.VehicleYear?.toString().toLowerCase().includes(term) ||
        x.VehicleMake?.toLowerCase().includes(term) ||
        x.VehicleModel?.toLowerCase().includes(term) ||

        x.DeliveryDate?.toString().toLowerCase().includes(term) ||
        x.DateSale?.toString().toLowerCase().includes(term) ||
        x.FundingDate?.toString().toLowerCase().includes(term) ||

        x.trade1year?.toString().toLowerCase().includes(term) ||
        x.trade1modelname?.toLowerCase().includes(term)

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

  get BalCITTotal(): number {
    return this.filteredData.reduce((total, item) => {
      return total + (parseFloat(item.BalCIT) || 0);
    }, 0);
  }

  get BalFloorplanTotal(): number {
    return this.filteredData.reduce((total, item) => {
      return total + (parseFloat(item.BalFloorplan) || 0);
    }, 0);
  }
  ////////////////////////////////////Pagination Code End////////////////////////////////////////////


  priorityAction: 'Y' | 'N' | null = null;
  selectedPriorityRecord: any = null;
  previousHasPriority: 'Y' | 'N' | null = null;
  priority(e: any, val: any, confirmtemplate: any, ref: any, refval: any) {
    this.previousHasPriority = val.HasPriority;
    console.log('previousHasPriority set to:', this.previousHasPriority);
    this.selectedPriorityRecord = val;
    if (ref == 'multi') {
      if (this.priorityRecords.length == 0) {
        alert('Please select atleast one record to prioritize');
        const element = <HTMLInputElement>document.getElementById('Priority');
        if (element) element.checked = false;
      } else {
        if (e.target.checked) {
          this.priorityAction = 'Y';
          this.priorityVisibility = true;
          this.ngbmodalActive = this.ngbmodal.open(confirmtemplate, { size: 'sm', backdrop: 'static' });
        }
      }
    } else {

      if (e.target.checked) {
        this.priorityAction = 'Y';
        this.priorityVisibility = true;
        this.priorityRecords.push(val);
      } else {
        this.priorityAction = 'N';

        // Remove from array if exists
        const index = this.priorityRecords.findIndex(
          (list: any) => list.StockNumner === refval
        );
        if (index > -1) {
          this.priorityRecords.splice(index, 1);
        }

        if (!e.target.checked && this.previousHasPriority == 'Y') {
          // Open confirmation modal
          this.ngbmodalActive = this.ngbmodal.open(confirmtemplate, {
            size: 'sm',
            backdrop: 'static'
          });
        }
      }
    }
  }
  confirmPriorityAction() {

    if (this.priorityAction === 'Y') {
      // ✅ Bulk / single PRIORITIZE
      this.priorityAdd();
      return;
    }

    // ❌ UNPRIORITIZE (single record)
    const payload = {
      schedulecontrolpriority: [
        {
          AS_ID: this.selectedPriorityRecord.storeid,
          Account: this.selectedPriorityRecord.Account,
          Scheduletype: 'CIT',
          Control: this.selectedPriorityRecord.Control,
          Stockno: this.selectedPriorityRecord.StockNumner,
          Dealno: this.selectedPriorityRecord.Deal,
          Custno: this.selectedPriorityRecord.CustomerNumber,
          UserID: this.userid,
          SetPriority: 'N'
        }
      ]
    };

    this.shared.api.postmethod('ReceivableExcludeControls/AddScheduleControlPriority', payload).subscribe(
      (res: any) => {
        if (res && res.status === 200) {
          alert('Priority removed successfully');
          (document.getElementById('closeadd') as HTMLInputElement)?.click();
          this.onclose()
          this.Getfloorplansdata();
          this.prioritycheck = false;
        } else {
          alert('Failed to unprioritize');
        }
      },
      () => {
        alert('502 Bad Gateway Error');
      }
    );
  }
  priorityAdd() {
    const payload = { schedulecontrolpriority: [] as any[] };

    for (let i = 0; i < this.priorityRecords.length; i++) {
      payload.schedulecontrolpriority.push({
        AS_ID: this.priorityRecords[i].storeid,
        Account: this.priorityRecords[i].Account,
        Scheduletype: "CIT",
        Control: this.priorityRecords[i].Control,
        Stockno: this.priorityRecords[i].StockNumner,
        Dealno: this.priorityRecords[i].Deal,
        Custno: this.priorityRecords[i].CustomerNumber,
        UserID: this.userid,
        SetPriority: "Y"
      });
    }
    if (payload.schedulecontrolpriority.length == 0) {
      alert('Please select atleast one record to prioritize');
      return;
    }
    this.shared.api.postmethod('ReceivableExcludeControls/AddScheduleControlPriority', payload).subscribe((res: any) => {
      if (res && res.status == 200) {
        alert('This control prioritized successfully');
        (document.getElementById('closeadd') as HTMLInputElement)?.click();
        this.onclose()
        this.Getfloorplansdata();
        this.prioritycheck = false;
        this.priorityRecords = [];
        this.priorityVisibility = false;
      } else {
        alert(res.status);
        try { this.shared.spinner.hide(); } catch { }
        this.NoData = true;
      }
    }, (error: any) => {
      alert('502 Bad Gate Way Error');
      try { this.shared.spinner.hide(); } catch { }
      this.NoData = true;
    });
  }

  currentCheckboxEvent: any = null;
  previousPriorityState: boolean = false;
  previousHeaderCheck: boolean = false;

  onPriorityCancel(modal: any) {
    this.FloorPlanData = JSON.parse(
      JSON.stringify(this.originalFloorPlanData)
    );

    this.prioritycheck = false
    this.priorityAction = null;
    this.selectedPriorityRecord = null;
    this.previousHasPriority = null;


    modal.dismiss();
  }

  // -------------------------------------
  // Merged report component methods (groups/stores/people)
  // -------------------------------------
  ngAfterViewInit(): void {
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.shared.common.pageName == 'CIT') {
        if (res.obj.storesData != undefined) {
          this.groupsArray = res.obj.storesData;
          this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
          this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_name : this.groupName = ''
          this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
          this.getStoresandGroupsValues()
        }
      }
    })
    // subscribe to reports and export triggers
    this.reportGetting = this.shared.api.GetReports().subscribe((data: any) => {
      if (data && data.obj && data.obj.Reference == 'CIT') {
        if (data.obj.header == undefined) {
          this.storeIds = data.obj.storeValues;
          this.vehiclear = data.obj.vehiclear;
          this.dealStatus = data.obj.dealStatus;
          this.saleType = data.obj.saleType;
          this.financeManagerId = data.obj.FIvalues;
          this.allordebit = data.obj.allordebit;

          this.groups = data.obj.groups;
          this.enablevehicle = this.vehiclear === 'WAR';
        } else {
          if (data.obj.header == 'Yes') {
            this.storeIds = data.obj.storeValues;
          }
        }

        if (this.storeIds != '') {
          this.Getfloorplansdata();
        } else {
          this.NoData = true;
          this.FloorPlanData = [];
        }

        const headerdata = {
          title: 'CIT',
          stores: this.storeIds,
          groups: this.groups,
          financemanagers: this.financeManagerId,
          dealStatus: this.dealStatus,
          saleType: this.saleType,
          allordebit: this.allordebit,
          vehiclear: this.vehiclear,
        };
        this.shared.api.SetHeaderData({ obj: headerdata });
        this.header = [{
          type: 'Bar',
          storeIds: this.storeIds,
          vehiclear: this.vehiclear,
          dealStatus: this.dealStatus,
          saleType: this.saleType,
          allordebit: this.allordebit,
          groups: this.groups,
          financemanagers: this.financeManagerId,
        }];
      }
    });

    this.excel = this.shared.api.getExportToExcelAllReports().subscribe((res: any) => {
      if (res && res.obj && res.obj.title == 'CIT' && res.obj.state == true) {
        //this.exportToExcel(); // merged export will create both sheets
      }
    });

    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res: any) => {
      if (res && res.obj && res.obj.title == 'CIT' && res.obj.statePrint == true) {
        this.printPDF();
      }
    });

    this.Pdf = this.shared.api.getExportToPDFAllReports().subscribe((res: any) => {
      if (res && res.obj && res.obj.title == 'CIT' && res.obj.statePDF == true) {
        this.generatePDF();
      }
    });

    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res: any) => {
      if (res && res.obj && res.obj.title == 'CIT' && res.obj.stateEmailPdf == true) {
        this.exportToEmailPDF(res.obj.Email, res.obj.notes, res.obj.from);
      }
    });
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
  AllorDebit(e: any) {
    this.allordebit = []
    // if (e == 'All') {
    //   if (this.dealStatus.length == 3) {
    //     this.dealStatus = []
    //     this.toast.warning('Please select atleast one dealStatus', '');

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
  StoresData(data: any) {
    this.storeIds = data.storeids;
    this.groupId = data.groupId;
    this.storename = data.storename;
    this.groupName = data.groupName;
    this.storecount = data.storecount;
    this.storedisplayname = data.storedisplayname;
  }
  ngOnDestroy() {
    if (this.reportGetting) this.reportGetting.unsubscribe();
    if (this.Pdf) this.Pdf.unsubscribe();
    if (this.print) this.print.unsubscribe();
    if (this.email) this.email.unsubscribe();
    if (this.excel) this.excel.unsubscribe();
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
    if (ref == 'ST') {
      const index = this.saleType.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.saleType.splice(index, 1);
      } else {
        this.saleType.push(e);
      }
    }
    if (ref == 'PT') {
      this.tempPriorityFilter = e;
    }
  }


  viewmoreAction(fp: any) {
    if (!fp.Notes) return;
    if (fp.Notesstate === '+') {
      fp.Notesstate = '-';
      fp.duplicateNotes = [...fp.Notes]; // full list (spread = new reference)
    } else {
      fp.Notesstate = '+';
      fp.duplicateNotes = fp.Notes.slice(0, 3); // first 3
    }
    this.cdRef.detectChanges();
  }
  viewDeal(dealData: any) {
    // const modalRef = this.ngbmodal.open(DealRecapComponent, { size: 'md', windowClass: 'compModal' });
    // modalRef.componentInstance.data = { dealno: dealData.Deal, storeid: dealData.storeid, stock: dealData.StockNumner, vin: dealData.vin, custno: dealData?.CustomerNumber }; // Pass data to the modal component    
    // modalRef.result.then((result) => {
    //   console.log(result); // Handle modal close result
    // }, (reason) => {
    //   console.log(`Dismissed: ${reason}`); // Handle dismiss reason
    // });
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

        } else {
          alert('Invalid Details');
        }
      },
      (error: any) => {
      },
    );
  }
  employees(block: any, e: any, ename?: any) {
    if (block === 'FM') {
      const index = this.selectedFiManagersvalues.findIndex((i: any) => i == e);

      if (index >= 0) {
        this.selectedFiManagersvalues.splice(index, 1);
      } else {
        this.selectedFiManagersvalues.push(e);
      }
      const index1 = this.selectedFiManagersvalues.findIndex((i: any) => i == e);
      if (index1 >= 0) {
        this.selectedFiManagersname = ename;
      }

      return;
    }
    if (block === 'AllFM') {
      if (e === 0) {
        this.selectedFiManagersvalues = this.financeManager.map((fm: any) => fm.FiId);
      } else if (e === 1) {
        this.selectedFiManagersvalues = [];
      }

      return;
    }
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

    this.dealStatus = this.dealStatus;
    this.saleType = this.saleType;
    ((this.groups = this.groups),
      (this.financeManagerId =
        this.selectedFiManagersvalues.length == this.financeManager.length
          ? '0'
          : this.selectedFiManagersvalues.toString()));
    this.Getfloorplansdata();
  }

  closeReportModal() {
    try { this.ngbmodal.dismissAll(); } catch { }
    this.Performance = 'Unload';
  }

  // ---------------------------------------------------------
  // Notes / hide logic (already in component)
  // ---------------------------------------------------------
  addNotes(item: any) {
    this.selecteddata = item;
    this.getDropDown(this.selecteddata.companyid);
    this.notesStageText = '';
    this.notesStageValue = '';
  }

  getDropDown(companyid: any) {
    const obj = {
      AssociatedReport: 'CIT',
      CompanyID: companyid
    };
    this.shared.api.postmethod(this.shared.common.routeEndpoint + 'GetScheduleNoteStages', obj).subscribe((res: any) => {
      if (res && res.status == 200) {
        this.notesstage = res.response;
        this.notesStageValue = '';
      }
    });
  }

  save() {
    if (this.notesStageText == '') {

      this.toast.show('Please enter notes', 'warning', 'Warning');
      return;
    }
    const obj = {
      AS_ID: this.selecteddata.storeid,
      Account: this.selecteddata.CIT_Account,
      Control: this.selecteddata.Control,
      Notes: this.notesStageText,
      StageId: this.notesStageValue,
      UserID: JSON.parse(localStorage.getItem('userInfo')!).userid
    };

    this.shared.api.postmethod(this.shared.common.routeEndpoint + 'AddScheduleNotesAction', obj).subscribe((res: any) => {
      if (res && res.status == 200) {
        try { this.shared.spinner.hide(); } catch { }

        this.toast.show('Notes Added Successfully', 'success', 'Success');
        this.callLoadingState = 'ANS';
        document.getElementById('close')?.click();
        this.onclose();
        this.commentsVisibility = true;

        const userName = JSON.parse(localStorage.getItem('userInfo')!)?.user_Info?.fullName;
        const curDate = new Date();
        let stageText = '';
        let nts = '';

        // ✅ Get Stage Text safely
        if (this.notesStageValue) {
          const filteredData = this.notesstage.filter((item: any) => item.NS_ID == this.notesStageValue);
          if (filteredData?.length) {
            stageText = filteredData[0].NS_Text.trim();
            this.selecteddata.Stage = stageText;
          }
        }

        // ✅ Build formatted note
        if (stageText) {
          nts = `${stageText} - ${this.notesStageText} - ${userName} - (${this.formatDate(curDate)})`;
        } else {
          nts = `${this.notesStageText} - ${userName} - (${this.formatDate(curDate)})`;
        }

        // ✅ Create note object
        const noteObj = {
          STAGE: stageText || '',
          NOTES: nts,
          NOTESDATE: this.formatDate(curDate),
          UserName: userName
        };

        // ✅ Update both Notes (full list) and duplicateNotes (display subset)
        if (!this.selecteddata.Notes) {
          this.selecteddata.Notes = [];
        }

        // Add new note at the top of full list
        this.selecteddata.Notes.unshift(noteObj);

        // Refresh the 3-note preview
        this.selecteddata.duplicateNotes = this.selecteddata.Notes.slice(0, 3);
        this.selecteddata.Notesstate = '+';
        this.selecteddata.NotesStatus = 'Y';

        this.cdRef.detectChanges(); // ✅ Force view refresh
      } else {

        this.toast.show('Something went wrong, please try again.', 'danger', 'Error');
      }
    });



  }

  collectHidevalues(e: any, val: any, confirmtemplate: any, ref: any, refval: any) {

    if (ref == 'multi') {
      if (this.hideRecords.length == 0) {

        this.toast.show('Please select atleast one record to hide', 'warning', 'Warning');
        const element = <HTMLInputElement>document.getElementById('symbol');
        if (element) element.checked = false;
      } else {
        if (e.target.checked) {
          this.hideVisibility = true;
          this.ngbmodalActive = this.ngbmodal.open(confirmtemplate, { size: 'sm', backdrop: 'static' });
        }
      }
    } else {

      if (e.target.checked) {
        this.hideVisibility = true;
        this.hideRecords.push(val);
      } else {
        const index = this.hideRecords.findIndex((list: { StockNumner: any }) => list.StockNumner == refval);
        this.hideRecords.splice(index, 1);
      }
    }
  }

  hideAdd() {

    this.FinalArray = [];
    for (let i = 0; i < this.hideRecords.length; i++) {
      const account = [this.hideRecords[i].CIT_Account, this.hideRecords[i].Vehicle_Account]
        .filter(Boolean)
        .join(', ');
      this.FinalArray.push({
        Receivable_Type: 'CIT',
        Account: account,
        CompanyID: this.hideRecords[i].companyid,
        AS_ID: this.hideRecords[i].storeid,
        Control: this.hideRecords[i].Control,
        Stock: this.hideRecords[i].StockNumner,
        Deal: this.hideRecords[i].Deal,
        Control_Status: 'Y'
      });
    }
    if (this.FinalArray.length == 0) {

      this.toast.show('Please select atleast one record to hide', 'warning', 'Warning');
      return;
    }
    const obj = { receivableexcludecontrol: this.FinalArray };
    this.shared.api.postmethod('ReceivableExcludeControls', obj).subscribe((res: any) => {
      if (res && res.status == 200) {

        this.toast.show('This control hidden successfully', 'success', 'Success');
        (document.getElementById('closeadd') as HTMLInputElement)?.click();
        this.onclose()
        this.Getfloorplansdata();
        this.hideRecords = [];
        this.hideVisibility = false;
      } else {

        this.toast.show(res.status, 'danger', 'Error');
        try { this.shared.spinner.hide(); } catch { }
        this.NoData = true;
      }
    }, (error: any) => {

      this.toast.show('502 Bad Gate Way Error', 'danger', 'Error');
      try { this.shared.spinner.hide(); } catch { }
      this.NoData = true;
    });
  }

  public inTheGreen(value: number): boolean {
    return value >= 0;
  }

  // ---------------------------------------------------------
  // UI helpers / scrolling / sorting / popovers
  // ---------------------------------------------------------
  updateVerticalScroll(event: any): void {
    this.scrollCurrentposition = event.target.scrollTop;
    const scrollDemo = document.querySelector('#scrollcent') as HTMLElement;
    this.Scrollpercent = Math.round(
      (event.target.scrollTop / (event.target.scrollHeight - (scrollDemo?.clientHeight || 1))) * 100
    );
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

  sort(property: any, state?: any) {
    if (state == undefined) {
      this.isDesc = !this.isDesc;
    }
    this.callLoadingState = 'FL';
    this.column = property;
    let direction = this.isDesc ? 1 : -1;
    this.FloorPlanData.sort(function (a: any, b: any) {
      if (a[property] < b[property]) {
        return -1 * direction;
      } else if (a[property] > b[property]) {
        return 1 * direction;
      } else {
        return 0;
      }
    });
  }

  openMenu(Object: any) { /* placeholder */ }

  notesView() {
    this.notesViewState = !this.notesViewState;
  }

  openComments() {
    this.commentsVisibility = !this.commentsVisibility;
  }
  index = '';
  clean(value: any) {
    if (Array.isArray(value)) {
      return value.join(', ');
    }
    if (typeof value === 'string' && value.startsWith('[')) {
      try {
        return JSON.parse(value).join(', ');
      } catch {
        return value.replace(/[\[\]"]+/g, '');
      }
    }
    return value || '-';
  }

// exportToExcelCIT() {
//   const workbook = new Workbook();
//   const worksheet = workbook.addWorksheet('CIT Report');

//   worksheet.views = [{
//     state: 'frozen',
//     ySplit: 15,
//     topLeftCell: 'A16',
//     showGridLines: false,
//   }];

//   // ---------------- TITLE
//   worksheet.addRow('');
//   const titleRow = worksheet.addRow(['CIT Report']);
//   titleRow.font = { name: 'Arial', size: 12, bold: true };
//   titleRow.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
//   worksheet.addRow('');

//   // ---------------- DATE
//   const DateToday = this.datepipe.transform(new Date(), 'MM.dd.yyyy h:mm:ss a');
//   const DATE_EXTENSION = this.datepipe.transform(new Date(), 'MMddyyyy');
//   worksheet.addRow([DateToday]).font = { name: 'Arial', size: 9 };

//   // ---------------- FILTERS
//   const ReportFilter = worksheet.addRow(['Report Controls :']);
//   ReportFilter.font = { name: 'Arial', size: 10, bold: true };

//   worksheet.getCell('A6').value = 'Group :';
//   worksheet.getCell('B6').value = this.groupName;
//   worksheet.getCell('B6').font = { name: 'Arial', size: 9 };
//   worksheet.getCell('B6').alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };

//   worksheet.mergeCells('B7:K9');
//   worksheet.getCell('A7').value = 'Stores :';
//   worksheet.getCell('B7').value = this.ExcelStoreNames
//     ? this.ExcelStoreNames.toString().replaceAll(',', ', ')
//     : 'All Stores';
//   worksheet.getCell('B7').font = { name: 'Arial', size: 9 };
//   worksheet.getCell('B7').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };

//   worksheet.addRow('');

//   // ---------------- TOTAL / AGING GRID HEADER
//   const TotalHeaderRow = worksheet.addRow(['', 'TOTAL', '0-5', '6-10', '11-15', '15+']);
//   TotalHeaderRow.eachCell(cell => {
//     cell.font = { name: 'Arial', size: 9, bold: true, color: { argb: 'FFFFFF' } };
//     cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '2a91f0' } };
//     cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
//     cell.alignment = { vertical: 'middle', horizontal: 'center' };
//   });

//   const dataRow = worksheet.addRow([
//     'CIT Aging',
//     this.FloorPlanTotalData[0].AgeData[0].TOTAL,
//     this.FloorPlanTotalData[0].AgeData[0].D1,
//     this.FloorPlanTotalData[0].AgeData[0].D2,
//     this.FloorPlanTotalData[0].AgeData[0].D3,
//     this.FloorPlanTotalData[0].AgeData[0].D4
//   ]);
//   dataRow.eachCell((cell, col) => {
//     if (col > 1) cell.numFmt = '$#,##0';
//     cell.font = { name: 'Arial', size: 9, bold: true, color: { argb: '000000' } };
//     cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
//     cell.alignment = { vertical: 'middle', horizontal: 'center' };
//   });

//   worksheet.addRow('');

//   // ---------------- COLUMN GROUP HEADERS
//   worksheet.mergeCells('F14:H14'); // Balances
//   worksheet.getCell('F14').value = 'Balances';
//   worksheet.mergeCells('I14:K14'); // Customer
//   worksheet.getCell('I14').value = 'Customer';
//   worksheet.mergeCells('L14:U14'); // Deal Info
//   worksheet.getCell('L14').value = 'Deal Info';
//   worksheet.mergeCells('V14:X14'); // Vehicle
//   worksheet.getCell('V14').value = 'Vehicle';
//   worksheet.mergeCells('Y14:Z14'); // Dates
//   worksheet.getCell('Y14').value = 'Dates';

//   ['F14','I14','L14','V14','Y14'].forEach(cell => {
//     const c = worksheet.getCell(cell);
//     c.font = { name: 'Arial', size: 9, bold: true, color: { argb: 'FFFFFF' } };
//     c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '2a91f0' } };
//     c.alignment = { horizontal: 'center', vertical: 'middle' };
//     c.border = { right: { style: 'thin' } };
//   });

//   // ---------------- COLUMN HEADINGS
//   const Headings = [
//     'Notes','Priority','CIT Age','Date','Account','Control','Control 2','CIT','Floorplan',
//     'Customer Name','Customer Number','Store','Stock #','Deal #','Stage','F&I Mgr','Sales Mgr',
//     'Sales Person','Type','New/Used','Status','Bank Name','Year','Make','Model','Delivery','Sale','Funding','Trade'
//   ];
//   const headerRow = worksheet.addRow(Headings);
//   headerRow.eachCell(cell => {
//     cell.font = { name: 'Arial', size: 8, bold: true, color: { argb: 'FFFFFF' } };
//     cell.alignment = { horizontal: 'center', vertical: 'middle' };
//     cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '788494' } };
//     cell.border = { right: { style: 'thin' } };
//   });

//   // ---------------- DATA ROWS + NOTES
//   let rowIndex = 15;
//   for (const d of this.FloorPlanData) {
//     rowIndex++;
//     const dataRow = worksheet.addRow([
//       '', d.HasPriority ?? '-', d.AGE ?? '-', d.AgeDate ?? '-', d.Account ?? '-', d.Control ?? '-', d.Control2 ?? '-',
//       d.BalCIT ?? 0, d.BalFloorplan ?? 0, d.CustomerName ?? '-', d.CustomerNumber ?? '-', d.store ?? '-',
//       d.StockNumner ?? '-', d.Deal ?? '-', d.Stage ?? '-', d.FIManager ?? '-', d.SalesManager ?? '-',
//       d.SP1 ?? '-', d.SaleType ?? '-', d.DealType ?? '-', d.DealStatus ?? '-', d.BankName ?? '-',
//       d.VehicleYear ?? '-', d.VehicleMake ?? '-', d.VehicleModel ?? '-', d.DeliveryDate ?? '-', d.DateSale ?? '-', d.FundingDate ?? '-', d.trade1modelname ?? '-'
//     ]);

//     dataRow.eachCell((cell, col) => {
//       if([8,9].includes(col)) cell.numFmt = '$#,##0';
//       cell.font = { name: 'Arial', size: 8 };
//       cell.alignment = { vertical: 'middle', horizontal: col > 9 ? 'left' : 'center' };
//       cell.border = { right: { style: 'thin' } };
//     });

//     // ---------- NOTES
//     if(d.Notes && d.Notes.length && this.commentsVisibility) {
//       for(const note of d.Notes) {
//         rowIndex++;
//         worksheet.mergeCells(rowIndex, 1, rowIndex, 28);
//         const noteCell = worksheet.getCell(rowIndex, 1);
//         noteCell.value = 'Notes: ' + (note?.NOTES ?? 'No additional notes.');
//         noteCell.font = { name: 'Arial', size: 9, italic: true };
//         noteCell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
//         noteCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F2F2F2' } };

//         if(note?.TS) {
//           rowIndex++;
//           worksheet.mergeCells(rowIndex, 1, rowIndex, 28);
//           const dateCell = worksheet.getCell(rowIndex, 1);
//           dateCell.value = '   ' + this.datepipe.transform(note.TS, 'MM.dd.yyyy');
//           dateCell.font = { name: 'Arial', size: 8 };
//           dateCell.alignment = { vertical: 'middle', horizontal: 'left' };
//           dateCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F2F2F2' } };
//         }
//       }
//     }
//   }

//   // ---------------- AUTO WIDTH
//   worksheet.columns.forEach((col:any) => {
//     let maxLength = 10;
//     col.eachCell({ includeEmpty: true }, (cell:any) => {
//       const len = cell.value ? cell.value.toString().length : 10;
//       if(len > maxLength) maxLength = len;
//     });
//     col.width = maxLength + 5;
//   });

//   // ---------------- SAVE FILE
//   workbook.xlsx.writeBuffer().then((data: any) => {
//     const blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
//     FileSaver.saveAs(blob, `CIT_Report_${DATE_EXTENSION}.xlsx`);
//   });
// }

  isNaN(value: any): boolean {
    return value !== null && value !== undefined && !isNaN(Number(value));
  }
  // placeholders for print/pdf/email - keep method names so triggers still work
  printPDF() {
    // copy or implement your print logic
  }

  generatePDF() {
    // copy or implement your pdf generation logic
  }

  exportToEmailPDF(email: any, notes: any, from: any) {
    // copy or implement your export-to-email logic
  }

  // utility
  formatDate(date: number | Date | undefined) {
    const options: Intl.DateTimeFormatOptions = {
      month: '2-digit' as '2-digit',
      day: '2-digit' as '2-digit',
      year: 'numeric' as 'numeric',
    };
    return new Intl.DateTimeFormat('en-US', options).format(date as Date);
  }
}
