

import { Component, ElementRef, ViewChild, HostListener, OnInit, EventEmitter, Output, Input, Renderer2 } from '@angular/core';
import { NgbActiveModal, NgbDateParserFormatter, NgbModal } from '@ng-bootstrap/ng-bootstrap';

import { Subscription } from 'rxjs';

import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { Setdates } from '../../../../Core/Providers/SetDates/setdates';
import { BsDatepickerConfig, BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { FilterPipe } from '../../../../Core/Providers/filterpipe/filter.pipe';
import { ChangeDetectorRef } from '@angular/core';
import { Stores } from '../../../../CommonFilters/stores/stores';
import { Notes } from '../../../../Layout/notes/notes';
import { DateRangePicker } from '../../../../CommonFilters/date-range-picker/date-range-picker';
import { BsDatepickerDirective } from 'ngx-bootstrap/datepicker';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';

const EXCEL_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
const EXCEL_EXTENSION = '.xlsx';

@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [SharedModule, BsDatepickerModule, FilterPipe, Notes, Stores],
  templateUrl: './dashboard.html',
  styleUrls: ['./dashboard.scss'],
})
export class Dashboard implements OnInit {
  @ViewChild(BsDatepickerDirective, { static: false })
  monthPicker!: BsDatepickerDirective;

  // Dates & filters (appointment-style)
  FromDate: any = '';
  ToDate: any = '';
  CurrentDate = new Date();


  minDate!: Date;
  maxDate!: Date;
  DateType: any = 'C';
  displaytime: any = '';
  // user / store
  storeIds: any = '';
  Role: any = [];
  userid: any;

  // UI state
  spinnerLoader: boolean = true;


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
  FinanceReserveReconData: any = [];
  FloorPlanTotalData: any = [];
  TotalFloorPlanData: any = [];
  QISearchName: any = '';
  commentdata: any = [];
  notesstage: any = [];



  // filter defaults
  dealStatus: any = ['InClosed', 'Sold'];


  // header & widgets
  header: any = [
    {
      type: 'Bar',
      storeIds: this.storeIds,

      dealStatus: this.dealStatus,

      groups: this.groups,

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

  // --- merged report-component members ---
  StoresIds: any = [];

  Performance: any = 'Load';

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
  bsRangeValue!: Date[];
  //  storeIds: any = '';
  storesFilterData: any = {
    'groupsArray': this.groupsArray, 'groupId': this.groupId, 'storesArray': this.stores, 'storeids': '1', 'type': 'M', 'others': 'N',
    'groupName': this.groupName, 'storename': this.storename, storecount: null, 'storedisplayname': this.storedisplayname
  };

  Dates: any = {
    'FromDate': this.FromDate, 'ToDate': this.ToDate, "MaxDate": this.maxDate, 'MinDate': this.minDate, 'DateType': this.DateType, 'DisplayTime': this.displaytime,
    // Types: [
    //   { 'code': 'MTD', 'name': 'MTD' },
    //   { 'code': 'QTD', 'name': 'QTD' },
    //   { 'code': 'YTD', 'name': 'YTD' },
    //   { 'code': 'PYTD', 'name': 'PYTD' },
    //   { 'code': 'LY', 'name': 'Last Year' },
    //   { 'code': 'LM', 'name': 'Last Month' },
    //   { 'code': 'PM', 'name': 'Same Month PY' },
    // ]
  }


  selectedMonth!: Date;

  monthPickerConfig: Partial<BsDatepickerConfig> = {
    dateInputFormat: 'MMMM YYYY',
    minMode: 'month',   // 🔑 month picker
    adaptivePosition: true,
    showWeekNumbers: false,
    maxDate: new Date()
  };
  viewstore: any = '';

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

    public formatter: NgbDateParserFormatter,
    private toast: ToastService,
    // public toast: ToastrService,

  ) {
    localStorage.setItem('time', 'C');
    if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
      // this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.ustores.split(',')
      // this.storeIds = '1,2';
      this.groupId = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Preferences
      this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids.split(',')
    }



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
      //   if (this.comm && this.comm.redirectionFrom && this.comm.redirectionFrom.flag != 'M') {
      //     this.dealStatus = this.dealStatus;
      //     this.financeManagerId = 0;
      //     this.allordebit = 'dr';
      //   }
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


    this.bsConfig = {
      dateInputFormat: 'MM/dd/yyyy',
      showWeekNumbers: false,
      adaptivePosition: true
    };
    // ✅ Set present month by default
    const todayone = new Date();
    this.selectedMonth = new Date(todayone.getFullYear(), todayone.getMonth(), 1);

    // 🔁 Apply month logic
    this.monthSelected(this.selectedMonth);

    // if (this.shared.common.groupsandstores.length > 0) {
    //   this.groupsArray = this.shared.common.groupsandstores.filter((val: any) => val.sg_id != this.shared.common.reconID);
    //   this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
    //   this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_Name : this.groupName = ''
    //   this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
    //   this.getStoresandGroupsValues()
    // }
    // set page title
    try {
      this.shared.setTitle(this.shared.common.titleName + '-Finance Reserve Recon');
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
    // this.FromDate = this.ToDate.toISOString().slice(0, 10);
    // this.ToDate = this.CurrentDate.toISOString().slice(0, 10);
    let today = new Date();
    const tomorrow = new Date();
    tomorrow.setDate(today.getDate());

    this.FromDate = this.shared.datePipe.transform(today, 'MM/dd/yyyy');
    this.ToDate = this.shared.datePipe.transform(tomorrow, 'MM/dd/yyyy');

    // set header data unless favorite
    if (localStorage.getItem('Fav') != 'Y') {
      const data = {
        title: 'Finance Reserve Recon',
        stores: this.storeIds,
        groups: this.groups,
        dealStatus: this.dealStatus,
        fromdate: this.FromDate,
        todate: this.ToDate,
      };
      try {
        this.shared.api.SetHeaderData({ obj: data });
      } catch { /* ignore */ }

      this.header = [
        {
          type: 'Bar',
          storeIds: this.storeIds,
          dealStatus: this.dealStatus,
          groups: this.groups,
          fromdate: this.FromDate,
          todate: this.ToDate,
        },
      ];

      // if (this.storeIds != '') {
      this.GetFinanceReserveReconData();
      this.setDates(this.DateType)
      this.getEmployees()

      // }
    }

    // listen window clicks for certain modal close behavior (from report component)
    this.renderer.listen('window', 'click', (e: Event) => {
      const TagName = e.target as HTMLButtonElement;
      if (TagName && TagName.className === 'd-block modal fade show modal-static') {
        try { this.ngbmodal.dismissAll(); } catch { }
      }
    });
  }

  monthSelected(date: Date) {
    if (!date) return;

    const year = date.getFullYear();
    const month = date.getMonth();

    // First day of month
    this.FromDate = new Date(year, month, 1);

    // Last day of month
    this.ToDate = new Date(year, month + 1, 0);

    this.DateType = 'C';

    // Display text
    this.displaytime = this.shared.datePipe.transform(
      this.FromDate,
      'MMMM yyyy'
    );

    // keep Dates object in sync
    this.Dates.FromDate = this.FromDate;
    this.Dates.ToDate = this.ToDate;
    this.Dates.DateType = this.DateType;
    this.Dates.DisplayTime = this.displaytime;
  }
  ngOnInit(): void {
    // datepicker config and other init tasks
    this.bsConfig = {
      dateInputFormat: 'MM/dd/yyyy',
      showWeekNumbers: false,
      adaptivePosition: true
    };
  }

  openMonthPicker() {
    if (this.monthPicker) {
      this.monthPicker.show();
    }
  }

  open() {
    this.Dates = { ...this.Dates }
    this.setDates(this.DateType)
  }
  initializeDates(type: any) {
    let dates: any = this.setdates.setDates(type)
    this.FromDate = dates[0];
    this.ToDate = dates[1];
    localStorage.setItem('time', type);
    this.setDates(type)
  }
  setDates(type: any) {
    this.DateType == 'C' ? this.displaytime = this.shared.datePipe.transform(this.FromDate, 'MM/dd/yyyy') + ' - ' + this.shared.datePipe.transform(this.ToDate, 'MM/dd/yyyy') :
      this.displaytime = 'Time Frame (' + this.Dates.Types.filter((val: any) => val.code == type)[0].name + ')';
    // this.maxDate = new Date();
    // this.minDate = new Date();
    // this.minDate.setFullYear(this.maxDate.getFullYear() - 3);
    // this.maxDate.setDate(this.maxDate.getDate());
    this.Dates.FromDate = this.FromDate;
    this.Dates.ToDate = this.ToDate;
    this.Dates.MinDate = this.minDate;
    this.Dates.MaxDate = this.maxDate;
    this.Dates.DateType = this.DateType;
    this.Dates.DisplayTime = this.displaytime;
    console.log(this.FromDate, this.ToDate, this.DateType, this.displaytime, '..............');
  }

  updatedDates(data: any) {
    // console.log(data);
    this.FromDate = data.FromDate;
    this.ToDate = data.ToDate;
    this.DateType = data.DateType;
    this.displaytime = data.DisplayTime
    // this.bsRangeValue = [new Date(this.FromDate), new Date(this.ToDate)];
    this.bsRangeValue = [this.FromDate, this.ToDate];
  }

  // Messaging (report -> parent)
  sendMessage() {
    this.messageEvent.emit(this.QISearchName);
  }

  /* ------------------------------
     Core: Get floor plan / CIT data
     ------------------------------ */
  GetFinanceReserveReconData() {
    this.NoData = false;
    this.FinanceReserveReconData = [];
    this.FloorPlanTotalData = [];
    try { this.shared.spinner.show(); } catch { /* ignore */ }

    const obj = {
      AS_IDs: this.storeIds.toString(),
      Date: this.shared.datePipe.transform(this.FromDate, 'MM/dd/yyyy')
      // DealStatus: this.dealStatus.toString(),
      // ValueType: this.allordebit && this.allordebit.toString ? (this.allordebit.toString() == 'all' ? '' : this.allordebit.toString()) : '',
      // UserID: 0,
      // Retail_Fleet: this.saleType.toString(),
    };

    this.shared.api.postmethod(this.shared.common.routeEndpoint + 'GetDealFinanceReserveReconciliation', obj).subscribe(
      (res: any) => {
        try { this.shared.spinner.hide(); } catch { }
        if (res && res.status == 200) {
          if (res.response && res.response.length > 0) {
            this.FinanceReserveReconData = res.response.filter((e: any) => e.store != 'TOTAL');
            this.FloorPlanTotalData = res.response.filter((e: any) => e.store == 'TOTAL');

            // Parse AgeData / Comments / Notes safely
            this.FinanceReserveReconData.forEach((x: any) => {
              if (x && x.AgeData && typeof x.AgeData === 'string') {
                try { x.AgeData = JSON.parse(x.AgeData); } catch { }
              }
              if (x && x.COMMENT && typeof x.COMMENT === 'string') {
                try { x.COMMENT = JSON.parse(x.COMMENT); } catch { }
              }
              if (x && x.History_FINANCE_RSV_PER_ROUTEONE && typeof x.History_FINANCE_RSV_PER_ROUTEONE === 'string') {
                try { x.History_FINANCE_RSV_PER_ROUTEONE = JSON.parse(x.History_FINANCE_RSV_PER_ROUTEONE); } catch { }

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

            this.NoData = this.FinanceReserveReconData.length === 0;
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

  formatDiff(value: number): string {
    if (value === 0) return '$0.00';

    const absValue = Math.abs(value).toFixed(2);

    if (value < 0) {
      return `($${absValue})`;   // ✅ brackets for negative
    }

    return `$${absValue}`;
  }




  enableEdit(fp: any) {
    fp.isEditing = true;
    fp.editValue = fp.FINANCE_RSV_PER_ROUTEONE;
  }

  cancelEdit(fp: any) {
    fp.isEditing = false;
    fp.editValue = null;
  }

  saveFinanceRouteOne(fp: any) {
    fp.FINANCE_RSV_PER_ROUTEONE = Number(fp.editValue);
    fp.isEditing = false;

    const obj = {
      AS_ID: fp.AS_ID,
      Deal: fp.DEAL_NUMBER,
      FINRSV_ROUTEONE: fp.FINANCE_RSV_PER_ROUTEONE,
      UserID: JSON.parse(localStorage.getItem('userInfo')!)?.user_Info?.userid,

    };

    this.shared.api.postmethod(this.shared.common.routeEndpoint + 'SetDealFinrsvPerRouteOne', obj).subscribe(
      (res: any) => {
        try { this.shared.spinner.hide(); } catch { }
        if (res && res.status == 200) {
          this.toast.show('Finance Reserve per Route One Updated Successfully', 'success', 'Success');
          this.GetFinanceReserveReconData();
        } else {
          this.NoData = true;
        }
      },
      (error: any) => {
        try { this.shared.spinner.hide(); } catch { }
        this.NoData = true;
      }
    );

    // 🔹 OPTIONAL: Call API here
    // this.saveRouteOneValue(fp);
  }

  // -------------------------------------
  // Merged report component methods (groups/stores/people)
  // -------------------------------------
  ngAfterViewInit(): void {
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.shared.common.pageName == 'Finance Reserve Recon') {
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
      if (data && data.obj && data.obj.Reference == 'Finance Reserve Recon') {
        if (data.obj.header == undefined) {
          this.storeIds = data.obj.storeValues;
          this.dealStatus = data.obj.dealStatus;
          this.FromDate = data.obj.fromdate;
          this.ToDate = data.obj.todate;
          this.groups = data.obj.groups;
        } else {
          if (data.obj.header == 'Yes') {
            this.storeIds = data.obj.storeValues;
          }
        }

        if (this.storeIds != '') {
          this.GetFinanceReserveReconData();
        } else {
          this.NoData = true;
          this.FinanceReserveReconData = [];
        }

        const headerdata = {
          title: 'Finance Reserve Recon',
          stores: this.storeIds,
          groups: this.groups,
          dealStatus: this.dealStatus,
          fromdate: this.FromDate,
          todate: this.ToDate,
        };
        this.shared.api.SetHeaderData({ obj: headerdata });
        this.header = [{
          type: 'Bar',
          storeIds: this.storeIds,
          dealStatus: this.dealStatus,
          groups: this.groups,
          fromdate: this.FromDate,
          todate: this.ToDate,
        }];
      }
    });

    this.excel = this.shared.api.getExportToExcelAllReports().subscribe((res: any) => {
      if (res && res.obj && res.obj.title == 'Finance Reserve Recon' && res.obj.state == true) {
        this.exportToExcel(); // merged export will create both sheets
      }
    });

    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res: any) => {
      if (res && res.obj && res.obj.title == 'Finance Reserve Recon' && res.obj.statePrint == true) {
        this.printPDF();
      }
    });

    this.Pdf = this.shared.api.getExportToPDFAllReports().subscribe((res: any) => {
      if (res && res.obj && res.obj.title == 'Finance Reserve Recon' && res.obj.statePDF == true) {
        this.generatePDF();
      }
    });

    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res: any) => {
      if (res && res.obj && res.obj.title == 'Finance Reserve Recon' && res.obj.stateEmailPdf == true) {
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

    this.cdRef.detectChanges(); // ensure immediate UI update
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
      AS_ID: 2,
      type: 'F',
    };
    this.shared.api.postmethod(this.shared.common.routeEndpoint + 'GetEmployeesDev', obj).subscribe(
      (res: any) => {
        if (res && res.status == 200) {
          // if (val == 'F') {
          this.financeManager = res.response.filter((e: any) => e.FiName != 'Unknown');
          this.selectedFiManagersvalues = this.financeManager.map(function (a: any) { return a.FiId; });

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
          this.toast.show('Invalid Details.', 'danger', 'Error');
        }
      },
      (error: any) => { /* ignore console errors */ }
    );
  }

  employees(block: any, e: any, ename?: any) {
    if (block == 'FM') {
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
    }
    if (block == 'AllFM') {
      if (this.selectedFiManagersvalues.length == this.financeManager.length) {
        this.selectedFiManagersvalues = [];
      } else {
        this.selectedFiManagersvalues = this.financeManager.map(function (a: any) { return a.FiId; });
      }
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
    this.FromDate = new Date(this.FromDate);
    this.ToDate = new Date(this.ToDate);

    this.GetFinanceReserveReconData();
  }

  closeReportModal() {
    try { this.ngbmodal.dismissAll(); } catch { }
    this.Performance = 'Unload';
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
    this.FinanceReserveReconData.sort(function (a: any, b: any) {
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


  notesData: any = {}
  Notespopup: any;
  scrollpositionstoring: any;
  addNotes(data: any, ref: any) {

    this.scrollpositionstoring = this.scrollCurrentposition;

    this.notesData = {
      store: data.AS_ID,
      title1: data.DEAL_NUMBER,
      title2: 'FRR',
      apiRoute: 'AddGeneralNotes',
      fulldata: data,

      // 🔥 IMPORTANT FIX
      history: Array.isArray(data.History_FINANCE_RSV_PER_ROUTEONE)
        ? data.History_FINANCE_RSV_PER_ROUTEONE
        : []
    };

    const hasHistory = this.notesData.history.length > 0;

    this.Notespopup = this.shared.ngbmodal.open(ref, {
      size: 'lg',
      backdrop: 'static',
      centered: true,
      windowClass: hasHistory ? 'notes-with-history' : 'notes-only'
    });

  }

  closeNotes(e: any) {
    this.shared.ngbmodal.dismissAll()
    if (e == 'S') {
      // this.details = []
      // this.cardealsdata = [];
      this.callLoadingState = 'ANS'
      this.GetFinanceReserveReconData()
    }
  }

  exportToExcel(): void {
    const workbook: any = this.shared.getWorkbook ? this.shared.getWorkbook() : null;
    if (!workbook) {
     
      this.toast.show('Workbook helper not available in shared service');
      return;
    }

    /* ================= Sheet ================= */
    const sheet = workbook.addWorksheet('Finance Reserve Recon');

    /* ================= Title ================= */
    const titleRow = sheet.addRow(['Finance Reserve Recon']);
    titleRow.font = { size: 14, bold: true };
    sheet.mergeCells('A1:G1');
    titleRow.alignment = { horizontal: 'center' };

    sheet.addRow([]);

    /* ================= Filters ================= */
    const filters = [
      { name: 'Store :', value: this.storedisplayname || 'All' },
      { name: 'Month :', value: this.selectedMonth ? this.selectedMonth.toLocaleDateString('en-US', { month: 'long', year: 'numeric' }) : '-' }
    ];

    filters.forEach(f => {
      const r = sheet.addRow([f.name, f.value]);
      r.getCell(1).font = { bold: true };
      sheet.mergeCells(`B${r.number}:G${r.number}`);
    });

    sheet.addRow([]);

    /* ================= Table Header ================= */
    const headers = [
      // 'Notes',
      'Store',
      'Finance Reserve',

      'Deal #',
      'Customer',
      'Lender',
      'Finance RSV Per Route One',
      'Diff'
    ];

    const headerRow = sheet.addRow(headers);
    headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };

    headerRow.eachCell((cell: { fill: { type: string; pattern: string; fgColor: { argb: string; }; }; alignment: { horizontal: string; vertical: string; }; }) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF2A91F0' }
      };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
    });

    /* ================= Data Rows ================= */
    (this.FinanceReserveReconData || []).forEach((fp: { COMMENT: any[]; STORE_NAME: any; FINANCE_RESERVE: any; DEAL_NUMBER: any; CUSTOMER: any; LENDER: any; FINANCE_RSV_PER_ROUTEONE: any; DIFF: number; IS_HIGHLIGHT: string; }) => {

      /* ----- Notes ----- */
      let notesText = '-';
      if (fp.COMMENT && fp.COMMENT.length > 0) {
        notesText = fp.COMMENT.map((n: any) => n.GN_TEXT).join(' | ');
      }

      const row = sheet.addRow([
        // notesText,
        fp.STORE_NAME ?? '-',

        fp.FINANCE_RESERVE ?? '-',
        fp.DEAL_NUMBER || '-',
        fp.CUSTOMER || '-',
        fp.LENDER || '-',
        fp.FINANCE_RSV_PER_ROUTEONE ?? '-',
        fp.DIFF ?? '-'
      ]);
      const diffCell = row.getCell('F'); // Diff column

      if (typeof fp.DIFF === 'number' && fp.DIFF < 0) {
        diffCell.font = {
          color: { argb: 'FFFF0000' }, // red
          bold: true
        };
      }
      /* ----- Currency formatting ----- */
      ['A', 'E', 'F'].forEach(col => {
        const cell = row.getCell(col);
        if (typeof cell.value === 'number') {
          cell.numFmt = '$#,##0.00;[Red]-$#,##0.00';
          cell.alignment = { horizontal: 'right' };
        }
      });

      /* ----- Highlight negative diff ----- */
      if (fp.DIFF < 0) {
        row.getCell('G').font = { color: { argb: 'FFFF0000' } };
      }

      /* ----- Highlight flagged rows ----- */
      if (fp.IS_HIGHLIGHT === 'Y') {
        row.eachCell((cell: { fill: { type: string; pattern: string; fgColor: { argb: string; }; }; }) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFFF4CC' }
          };
        });
      }
    });

    /* ================= Column Widths ================= */
    sheet.columns = [
      // { width: 35 }, // Notes
      { width: 18 }, // Finance Reserve
      { width: 12 }, // Deal #
      { width: 25 }, // Customer
      { width: 25 }, // Lender
      { width: 28 }, // RouteOne
      { width: 15 }  // Diff
    ];

    /* ================= Export ================= */
    workbook.xlsx.writeBuffer().then(() => {
      this.shared.exportToExcel(workbook, 'Finance_Reserve_Recon');
    }).catch((err: any) => {
      console.error(err);
    
      this.toast.show('Excel export failed', 'warning', 'Warning');
    });
  }


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
