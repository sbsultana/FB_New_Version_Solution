

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
  imports: [SharedModule, BsDatepickerModule, FilterPipe, TimeConversionPipe, Stores],
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
  dealStatus: any = ['Booked', 'Finalized'];
  saleType: any = ['Retail', 'Fleet'];
  allordebit: any = 'dr';
  financeManagerId: any = '0';

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

  // --- merged report-component members ---
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

      // if (this.storeIds != '') {
      this.Getfloorplansdata();
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
    try { this.shared.spinner.show(); } catch { /* ignore */ }

    const obj = {
      AS_ID: this.storeIds.toString(),
      DealStatus: this.dealStatus.toString(),
      ValueType: this.allordebit && this.allordebit.toString ? (this.allordebit.toString() == 'all' ? '' : this.allordebit.toString()) : '',
      FIManagerID: this.financeManagerId,
      UserID: 0,
      Aging_From: this.agingFilter?.min ?? null,
      Aging_To: this.agingFilter?.max ?? null,
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
        this.exportToExcel(); // merged export will create both sheets
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

  // Groups & Stores logic (merged from report)
  // getGroups() {
  //   if (this.comm && this.comm.groupsandstores) {
  //     if (this.comm.groupsandstores.length > 0) {
  //       this.groups = this.comm.groupsandstores.filter((val: any) => val.sg_id != '4');
  //       this.groupsBindingData();
  //     }
  //   }
  // }

  // groupsBindingData() {
  //   this.service.GetHeaderData().subscribe((res: any) => {
  //     if (res && res.obj && res.obj.title == 'CIT' && this.Performance == 'Load') {
  //       if (res.obj.groups != '' && res.obj.groups != '0') {
  //         this.selectedGroups = [];
  //         if (res.obj.groups.toString().indexOf(',') > 0) {
  //           var result = res.obj.groups.split(',');
  //           this.selectedGroups = result.map(function (x: any) { return parseInt(x, 10); });
  //           this.AllGroups = this.selectedGroups.length == this.groups.length ? true : false;
  //         } else {
  //           this.selectedGroups.push(parseInt(res.obj.groups));
  //           this.AllGroups = this.selectedGroups.length == this.groups.length ? true : false;
  //         }
  //         this.getGroupBaseStores(this.selectedGroups.toString());
  //       }
  //       if (res.obj.groups == '' || res.obj.groups == '0') {
  //         this.selectedGroups = this.groups.map(function (a: any) { return a.sg_id; });
  //         this.AllGroups = this.selectedGroups.length == this.groups.length ? true : false;
  //         this.getGroupBaseStores(this.selectedGroups.toString());
  //       }
  //     }
  //   });
  // }

  // getGroupBaseStores(id: any) {
  //   this.spinnerLoader = true;
  //   if (!this.comm || !this.comm.groupsandstores) {
  //     this.stores = [];
  //     this.spinnerLoader = false;
  //     return;
  //   }
  //   this.stores = this.comm.groupsandstores.filter((v: any) => v.sg_id == id)[0]?.Stores || [];
  //   this.spinnerLoader = false;
  //   this.selectedstorevalues = this.stores.map((a: any) => a.ID);
  //   if (this.groupstate == true) {
  //     this.getEmployees('F', this.selectedstorevalues.toString(), '2', 'Bar');
  //   }
  //   this.AllStores = true;
  //   this.service.GetHeaderData().subscribe((res: any) => {
  //     if (res && res.obj && res.obj.title == 'CIT' && this.Performance == 'Load' && this.groupstate == false) {
  //       if (res.obj.stores != '' && res.obj.stores != '0') {
  //         this.selectedstorevalues = [];
  //         var result = res.obj.stores.split(',');
  //         this.selectedstorevalues = result.map(function (x: any) { return parseInt(x, 10); });
  //         this.AllStores = this.selectedstorevalues.length == this.stores.length ? true : false;
  //       }
  //       if (res.obj.stores == '0') {
  //         this.selectedstorevalues = this.stores.map(function (a: any) { return a.ID; });
  //         this.AllStores = this.selectedstorevalues.length == this.stores.length ? true : false;
  //       }
  //       if (res.obj.stores == '') {
  //         this.selectedstorevalues = [];
  //         this.AllStores = false;
  //       }
  //       if (this.selectedstorevalues.length == 1 && this.stores.length > 0) {
  //         this.storeName = this.stores.filter((val: any) => val.ID == this.selectedstorevalues.toString())[0]?.storename;
  //       }
  //       this.groupName = this.groups.filter((val: any) => val.sg_id == this.selectedGroups[0])[0]?.sg_name;
  //     }
  //   });
  // }

  // selectedstoreToggle(e: any) {
  //   const index = this.selectedstorevalues.findIndex((i: any) => i == e.ID);
  //   if (index >= 0) {
  //     this.selectedstorevalues.splice(index, 1);
  //     this.AllStores = false;
  //   } else {
  //     this.selectedstorevalues.push(e.ID);
  //     this.AllStores = (this.selectedstorevalues.length === this.stores.length);
  //   }
  //   if (this.selectedstorevalues.length == 1) {
  //     this.storeName = this.stores.filter((val: any) => val.ID == this.selectedstorevalues.toString())[0]?.storename;
  //   }
  //   this.getEmployees('F', this.selectedstorevalues.toString(), '2', '');
  // }

  // allstores() {
  //   this.AllStores = !this.AllStores;
  //   if (this.AllStores == true) {
  //     this.selectedstorevalues = this.stores.map(function (a: any) { return a.ID; });
  //   } else {
  //     this.selectedstorevalues = [];
  //   }
  //   this.getEmployees('F', this.selectedstorevalues.toString(), '2', '');
  // }

  // allgroups() {
  //   this.groupstate = true;
  //   this.AllGroups = !this.AllGroups;
  //   if (this.AllGroups == true) {
  //     this.selectedGroups = this.groups.map(function (a: any) { return a.sg_id; });
  //     this.getGroupBaseStores(this.selectedGroups);
  //   } else {
  //     this.selectedGroups = [];
  //     this.stores = [];
  //     this.selectedstorevalues = [];
  //     this.AllStores = false;
  //   }
  // }

  // individualgroups(e: any) {
  //   this.groupstate = true;
  //   const index = this.selectedGroups.findIndex((i: any) => i == e.sg_id);
  //   if (index >= 0) {
  //     // ignore remove per original logic
  //   } else {
  //     this.selectedGroups = [];
  //     this.selectedGroups.push(e.sg_id);
  //     this.groupName = this.groups.filter((val: any) => val.sg_id == this.selectedGroups[0])[0]?.sg_name;
  //     this.getGroupBaseStores(this.selectedGroups.toString());
  //     this.AllGroups = (this.selectedGroups.length === this.groups.length);
  //   }
  // }

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
   
          this.toast.show('Invalid Details', 'danger', 'Error');
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
    this.saleType = this.saleType;
    this.financeManagerId = this.selectedFiManagersvalues.length == this.financeManager.length
      ? '0'
      : this.selectedFiManagersvalues.toString()
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

        const userName =JSON.parse(localStorage.getItem('userInfo')!)?.user_Info?.fullName;
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

  exportToExcel(): void {
    // Uses shared.getWorkbook() (as in your appointment example)
    // Creates two sheets:
    // 1) 'Dashboard Summary' - simple summary metrics
    // 2) 'Floorplan Details' - row-per-record of FloorPlanData
    const workbook: any = this.shared.getWorkbook ? this.shared.getWorkbook() : null;
    if (!workbook) {
     
      this.toast.show('Workbook helper not available in shared service', 'danger', 'Error');
      return;
    }

    // --- Dashboard Summary sheet ---
    const summarySheet = workbook.addWorksheet('Dashboard Summary');
    summarySheet.addRow(['CIT']);
    let storeValue = '';

    if (!this.storeIds || this.storeIds.length === 0) {
      // No selection → All stores
      storeValue = this.stores.map((s: any) => s.storename).join(', ');
    }
    else if (this.storeIds.length === this.stores.length) {
      // All selected → bind all store names
      storeValue = this.stores.map((s: any) => s.storename).join(', ');
    }
    else {
      // Partial selection
      storeValue = this.stores
        .filter((s: any) => this.storeIds.includes(s.ID))
        .map((s: any) => s.storename)
        .join(', ');
    }

    const filters = [
      { name: 'Store:', values: storeValue },
      { name: 'Deal Status :', values: this.clean(this.dealStatus).replace('Finalized', 'Closed or Sold') },
      { name: 'Sale Type :', values: this.clean(this.saleType) }
    ]
    // const ReportFilter = summarySheet.addRow(['Report Controls :']);
    // ReportFilter.font = { name: 'Arial', family: 4, size: 10, bold: true };
    let startIndex = 2
    filters.forEach((val: any) => {
      startIndex++
      summarySheet.addRow('');
      summarySheet.getCell(`A${startIndex}`);
      summarySheet.mergeCells(`B${startIndex}:C${startIndex}`);
      summarySheet.getCell(`A${startIndex}`).value = val.name;
      summarySheet.getCell(`B${startIndex}`).value = val.values
    })
    summarySheet.addRow('');

    // summarySheet.addRow([ '', 'Total', '0-5', '6-10', '11-15', '15+']).font = { bold: true };

    const agingHeader = summarySheet.addRow(['', 'Total', '0-5', '6-10', '11-15', '15+']);
    agingHeader.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    agingHeader.eachCell((cell: any) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF2A91F0' }
      };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
    });

    // ---- Aging Values Row ----
    if (this.FloorPlanData?.length > 0 && this.FloorPlanData[0]?.AgeData?.length > 0) {
      const A = this.FloorPlanData[0].AgeData[0];

      const agingValueRow = summarySheet.addRow([

        'CIT Aging',
        A.TOTAL ?? '-',
        A.D1 ?? '-',
        A.D2 ?? '-',
        A.D3 ?? '-',
        A.D4 ?? '-'
      ]);
      [2, 3, 4, 5, 6].forEach(col => {
        const cell = agingValueRow.getCell(col);
        cell.numFmt = '"$"#,##0.00;[Red]-"$"#,##0.00';
        cell.alignment = { horizontal: 'right' };
      });
    }

    summarySheet.addRow([]);

    const headers = [
      'CIT Age', 'Date', 'Control', 'Control 2', 'CIT Acct',
      'CIT', '	Vehicle Acct	', 'Vehicle', 'Balance', 'Customer', 'Number	', 'Store', 'Stock #', '	Deal #', 'Stage', 'F&I Mgr', 'Sales Mgr',
      'Sales Person', 'New/Used', 'Type', 'Status', 'Bank Name', 'Year', 'Make', 'Model', 'Delivery', 'Sale', 'Funding', 'Trade'
    ];
    // summarySheet.addRow(headers);
    const headerRow = summarySheet.addRow(headers);
    headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };

    headerRow.eachCell((cell: any) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF2A91F0' } // Blue header
      };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
    });
    (this.FloorPlanData || []).forEach((r: any) => {
      // ensure date fields formatted: null/empty -> '-'
      const formatDateSafe = (d: any) => {
        if (!d) return '-';
        try {
          const dt = new Date(d);
          if (isNaN(dt.getTime())) return d?.toString() || '-';
          // use MM.dd.yyyy as in your template request
          const mm = String(dt.getMonth() + 1).padStart(2, '0');
          const dd = String(dt.getDate()).padStart(2, '0');
          const yyyy = dt.getFullYear();
          return `${mm}.${dd}.${yyyy}`;
        } catch {
          return d?.toString() || '-';
        }
      };
      let tradeYear = r.trade1year ? r.trade1year : '';
      let tradeModel = r.trade1modelname ? r.trade1modelname : '';

      let tradeValue = (tradeYear + ' ' + tradeModel).trim();

      if (!tradeValue) tradeValue = '-';
      const formatDateMMDDYYYY = (d: any) => {
        if (!d) return '-';

        const dt = new Date(d);
        if (isNaN(dt.getTime())) return '-';

        const mm = String(dt.getMonth() + 1).padStart(2, '0');
        const dd = String(dt.getDate()).padStart(2, '0');
        const yyyy = dt.getFullYear();

        return `${mm}/${dd}/${yyyy}`;
      };

      const row = [
        r.AGE || '-',
        formatDateMMDDYYYY(r.AgeDate || '-'),
        // r.CIT_Account || '-',
        r.Control || '-',
        r.Control2 || '-',
        r.CIT_Account || '-',
        r.CIT_Balance || '-',
        r.Vehicle_Account || '-',
        r.Vehicle_Balance || '-',
        r.CIT_Vehicle_Balance || '_',
        // r.BalFloorplan || '-',
        r.CustomerName || '-',
        r.CustomerNumber || '-',
        r.store || '-',
        r.StockNumner || '-',
        r.Deal || '-',
        r.Stage || '-',
        r.FIManager || '-',
        r.SalesManager || '-',
        r.SP1 || '-',
        r.SaleType || '-',
        r.DealType || '-',
        r.DealStatus || '-',
        r.BankName || '-',
        r.VehicleYear || '-',
        r.VehicleMake || '-',
        r.VehicleModel || '-',
        r.DeliveryDate || '-',
        r.DateSale || '-',
        r.FundingDate || '-',
        tradeValue,


      ];
      summarySheet.addRow(row);

      if (r.duplicateNotes && r.duplicateNotes.length > 0) {

        summarySheet.addRow([]); // spacing before notes

        const notesHeader = summarySheet.addRow(["Notes"]);
        notesHeader.font = { bold: true };

        r.duplicateNotes.forEach((n: any) => {
          summarySheet.addRow(["" + (n.NOTES || "-")]);
        });
      }
    });
    const currencyColumns = [6, 8, 9];

    currencyColumns.forEach(colIndex => {
      const column = summarySheet.getColumn(colIndex);

      column.numFmt = '"$"#,##0.00;[Red]-"$"#,##0.00';
      column.alignment = { horizontal: 'right' };
    });
    // set some column widths
    summarySheet.columns.forEach((col: any) => {
      col.width = 20;
    });
    // ================= TOTALS ROW =================
    if (this.FloorPlanTotalData?.length > 0) {
      const T = this.FloorPlanTotalData[0];

      const totalsRow = summarySheet.addRow([
        '',                 // CIT Age
        '',                 // Date
        '',                 // Control
        '',                 // Control 2
        'Totals',           // Label
        // '',                 // CIT Acct
        T.CIT_Balance || 0, // CIT Balance
        '',                 // Vehicle Acct
        T.Vehicle_Balance || 0,
        T.CIT_Vehicle_Balance || 0,
        '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''
      ]);

      // ===== Styling =====
      totalsRow.font = { bold: true };

      totalsRow.eachCell((cell: any) => {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFEFEFEF' } // light gray
        };
      });

      // ===== Currency formatting =====
      [7, 9, 10].forEach(col => {
        totalsRow.getCell(col).numFmt = '"$"#,##0.00;[Red]-"$"#,##0.00';
        totalsRow.getCell(col).alignment = { horizontal: 'right' };
      });
    }

    // write workbook to buffer and trigger shared export helper
    workbook.xlsx.writeBuffer().then((buffer: any) => {
      try {
        // if shared has exportToExcel helper (appointment example), use it
        if (this.shared.exportToExcel) {
          this.shared.exportToExcel(workbook, 'CIT');
        } else {
          // fallback: use FileSaver if available globally (not implemented here)
        
          this.toast.show('Export helper not found. Implement shared.exportToExcel or add FileSaver logic.', 'danger', 'Error');

        }
      } catch (err) {
        console.error('Export error', err);
      
        this.toast.show('Export failed. See console for details.', 'danger', 'Error');
      }
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
