import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { Component, ViewChild, ElementRef, HostListener, SimpleChanges } from '@angular/core';
// import { Subscription } from 'rxjs';
import { Setdates } from '../../../../Core/Providers/SetDates/setdates';
import { BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { common } from '../../../../common';
import { Stores } from '../../../../CommonFilters/stores/stores';
import { Subscription } from 'rxjs';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
@Component({
  selector: 'app-dashboard',
  imports: [SharedModule, BsDatepickerModule, Stores],
  standalone: true,
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss'
})
export class Dashboard {
  SalesPersonsData: any = [];
  FIManagerData: any = [];
  TotalSalesPersonsData: any = [];
  FromDate: any;
  ToDate: any;
  NoData: boolean = false;
  path1: any = '';
  path2: any = '';
  path3: any = '';
  TotalReport: any = 'T';
  storeIds: any = '0';
  CompleteComponentState: boolean = true;
  dateType: any = 'MTD';
  dealType: any = 'New,Used';
  saleType: any = 'Retail,Lease,Misc,Special Order';
  financeType: any = 'Finance,Cash,Lease';
  storeorgrp: any = 'G';
  groups: any = 0;

  stores: any = []
  groupsArray: any = [];
  storename: any = ''
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  groupId: any = 0;
  otherStoresArray: any = [];
  otherStoreIds: any = [];
  excel!: Subscription;

  storesFilterData: any = {
    'groupsArray': this.groupsArray, 'groupId': this.groupId, 'storesArray': this.stores, 'storeids': '1', 'type': 'M', 'others': 'Y',
    'groupName': this.groupName, 'storename': this.storename, storecount: null, 'storedisplayname': this.storedisplayname,
    otherStoresArray: this.otherStoresArray, otherStoreIds: this.otherStoreIds

  };


  header: any = [{
    type: 'Bar', storeIds: this.storeIds, storeorgroup: this.storeorgrp, dealType: this.dealType,
    saleType: this.saleType,
    financeType: this.financeType, groups: this.groups
  }]
  popup: any = [{ type: 'Popup' }];

  // solutionurl: any = this.shared.getEnviUrl();
  LogCount = 1;





  columnName: any = 'Rank';
  columnState: any = 'asc';
  dealStatus: any;

  constructor(
    public shared: Sharedservice, public setdates: Setdates, private comm: common, private toast: ToastService,
  ) {
    // this.initializeDates();
    this.shared.setTitle(this.shared.common.titleName + 'F&I Manager Ranking');

    this.shared.setTitle(this.shared.common.titleName + '-F&I Manager Ranking');
    if (typeof window !== 'undefined') {
      if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {

        if (localStorage.getItem('flag') == 'V') {
          this.storeIds = [];
          console.log(JSON.parse(localStorage.getItem('userInfo')!), JSON.parse(localStorage.getItem('userInfo')!).user_Info, 'Widget Stores............');
          this.groupId = JSON.parse(localStorage.getItem('userInfo')!).groupid
          JSON.parse(localStorage.getItem('userInfo')!).store.indexOf(',') > 0 ?
            this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).store.split(',') :
            this.storeIds.push(JSON.parse(localStorage.getItem('userInfo')!).store)
          localStorage.setItem('flag', 'M')
        } else {
          this.groupId = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Preferences
          this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids.split(',');
          this.otherStoreIds = JSON.parse(localStorage.getItem('otherstoreids')!);

        }
      }
      if (this.shared.common.groupsandstores.length > 0) {
        this.groupsArray = this.shared.common.groupsandstores.filter((val: any) => val.sg_id != this.shared.common.reconID);
        this.otherStoresArray = this.shared.common.OtherStoresData[0].Stores

        this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
        this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_Name : this.groupName = ''
        this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
        this.getStoresandGroupsValues()
      }

      if (localStorage.getItem('stime') != null) {
        let stime = localStorage.getItem('stime');
        if (stime != null && stime != '') {
          this.initializeDates(stime)
          this.DateType = stime
        }
      } else {
        this.initializeDates('MTD')
        this.DateType = 'MTD'
      }
      // localStorage.setItem('stime', 'MTD')

    }
    this.shared.setTitle(this.shared.common.titleName + '-F&I Manager Ranking');
    if (typeof window !== 'undefined') {

      // if (localStorage.getItem('Fav') != 'Y') {
      const data = {
        title: 'F&I Manager Rankings',
        path1: '',
        path2: '',
        path3: '',
        stores: this.storeIds,
        toporbottom: this.TotalReport,
        datetype: '',
        fromdate: this.FromDate,
        todate: this.ToDate,
        groups: this.groups,
        storeorgroup: this.storeorgrp,
        dealType: this.dealType,
        saleType: this.saleType,
        financeType: this.financeType,
        otherstoreids: this.otherStoreIds,
        count: 0
      };
      this.shared.api.SetHeaderData({
        obj: data,
      });
      this.header = [{
        type: 'Bar', storeIds: this.storeIds, storeorgroup: this.storeorgrp, dealType: this.dealType,
        saleType: this.saleType,
        financeType: this.financeType, groups: this.groups, otherstoreids: this.otherStoreIds
      }]

      this.GetData('Rank', 'asc');
      // } else {
      //   this.getFavReports();
      // }
    }
  }

  ngOnInit(): void {
    this.SetDates(this.DateType);

  }



  GetData(sortdata?: any, sortstate?: any) {
    this.FIManagerData = [];
    this.shared.spinner.show();
    const obj = {
      StartDate: this.FromDate,
      EndDate: this.ToDate,
      StoreID: [...this.storeIds, ...this.otherStoreIds],
      Exp: sortdata,
      OrderType: sortstate,
      RankBy: this.storeorgroup,
      UserID: 0,
      SaleType: this.neworused.toString(),
      DealType: this.retailorlease.toString(),
      FinanceType: this.financetype.toString(),
    };
    let startFrom = new Date().getTime();
    const curl = this.shared.getEnviUrl() + 'cavender/GetFIManagerRankings';
    this.shared.api.postmethod(this.shared.common.routeEndpoint + 'GetFIManagerRankings', obj).subscribe(
      (res: { message: any; status: number; response: string | any[] | undefined; }) => {
        const currentTitle = document.title;
        this.shared.api.logSaving(curl, {}, '', res.message, currentTitle);
        if (res.status == 200) {
          if (res.response != undefined) {
            if (res.response.length > 0) {
              let resTime = (new Date().getTime() - startFrom) / 1000;
              // this.logSaving(
              //   this.solutionurl + 'cavender/GetFIManagerRankings',
              //   obj,
              //   resTime,
              //   'Success'
              // );
              this.FIManagerData = res.response;
              this.shared.spinner.hide();
              this.NoData = false;
              // this.FIManagerData.some(function (x: any) {
              //   x.Data2 = JSON.parse(x.Data2);
              //   x.Dealerx = '+';
              //   return false;
              // });
              // this.GetTotalData();
              let position = this.scrollCurrentposition + 5
              setTimeout(() => {
                if (this.scrollcent)
                  this.scrollcent.nativeElement.scrollTop = position
                // //console.log(position);

              }, 500);
            } else {
              // this.toast.error('Empty Response', '');
              this.shared.spinner.hide();
              this.NoData = true;
            }
          } else {
            let resTime = (new Date().getTime() - startFrom) / 1000;
            // this.logSaving(
            //   this.solutionurl + 'cavender/GetFIManagerRankings',
            //   obj,
            //   resTime,
            //   'Error'
            // );
            // this.shared.toast.error(res.status, '');

            this.shared.spinner.hide();
            this.NoData = true;
          }
        } else {
          // this.toast.error(res.status, '');
          this.shared.spinner.hide();
          this.NoData = true;
        }
      },
      (error: any) => {
        // this.toast.error('502 Bad Gate Way Error', '');
        this.shared.spinner.hide();
        this.NoData = true;
      }
    );
  }
  openDetails(data: any) {
    const DetailsSalesPeron = this.shared.ngbmodal.open(
      // SalesgrossDetailsComponent,
      {
        size: 'xxl',
        backdrop: 'static',
      }
    );
    DetailsSalesPeron.componentInstance.Salesdetails = [
      {
        StartDate: this.FromDate,
        EndDate: this.ToDate,
        dealtype: this.dealType,
        saletype: this.saleType,
        dealstatus: [
          "Delivered",
          "Capped",
          "Finalized"
        ],
        var1: 'store',
        var2: 'fimanager',
        var3: '',
        var1Value: data.StoreName,
        var2Value: data.FITrimName,
        var3Value: '',
        userName: data.FITrimName,
        FinanceManager: "0",
        SalesManager: "0",
        SalesPerson: "0"
      },
    ];
    DetailsSalesPeron.result.then(
      (data: any) => { },
      (reason: any) => {
        // on dismiss
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

  isDesc: boolean = false;
  column: string = 'CategoryName';
  sort(property: any) {
    this.isDesc = !this.isDesc; //change the direction
    this.column = property;
    let direction = this.isDesc ? 1 : -1;
    if (this.TotalReport == 'T') {
      var arr = this.FIManagerData.slice(1, this.FIManagerData.length);
      arr.sort(function (a: any, b: any) {
        if (a[property] < b[property]) {
          return -1 * direction;
        } else if (a[property] > b[property]) {
          return 1 * direction;
        } else {
          return 0;
        }
      });
      arr.unshift(this.FIManagerData[0]);
      this.FIManagerData = arr;
    } else {
      var arr = this.FIManagerData.slice(0, -1);
      arr.sort(function (a: any, b: any) {
        if (a[property] < b[property]) {
          return -1 * direction;
        } else if (a[property] > b[property]) {
          return 1 * direction;
        } else {
          return 0;
        }
      });
      arr.push(this.FIManagerData[this.FIManagerData.length - 1]);
      this.FIManagerData = arr;
    }
  }
  FIMstate: any;
  tabClick(col_Name: any, Col_state: any) {
    if (this.columnName == col_Name) {
      if (Col_state == 'asc') {
        this.columnState = 'desc';
        this.GetData(this.columnName, this.columnState);
      } else {
        this.columnState = 'asc';
        this.GetData(this.columnName, this.columnState);
      }
    } else {
      if (this.storeorgrp == 'G' && (col_Name != 'Rank' && col_Name != 'fimanager' && col_Name != 'StoreName')) {
        this.columnState = 'desc';
        this.columnName = col_Name;
        this.GetData(this.columnName, this.columnState);
      } else {
        this.columnState = 'asc';
        this.columnName = col_Name;
        this.GetData(this.columnName, this.columnState);
      }

    }
    // //console.log(this.columnName, this.columnState);
  }

  ngChanges: any = []
  ngOnChanges(changes: SimpleChanges) {
    console.log(changes, 'Report');
    this.ngChanges = changes['header'].currentValue[0];
    this.FromDate = this.ngChanges.fromDate;
    this.ToDate = this.ngChanges.toDate;
    this.neworused = this.ngChanges.dealType;
    this.retailorlease = this.ngChanges.saleType;
    // this.dealstatus = this.ngChanges.dealStatus;
    // this.setDates(this.ngChanges.datevaluetype)

  }
  ngAfterViewInit(): void {
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.shared.common.pageName == 'F&I Manager Rankings') {
        if (res.obj.storesData != undefined) {
          this.groupsArray = res.obj.storesData;
          this.otherStoresArray = this.shared.common.OtherStoresData[0].Stores

          this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
          this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_name : this.groupName = ''
          this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
          this.getStoresandGroupsValues()
        }
      }
    })
    this.shared.api.GetReportOpening().subscribe((res: { obj: { Module: string; }; }) => {
      // //console.log(res);
      if (res.obj.Module == 'F&I Manager Rankings') {
        document.getElementById('report')?.click()
      }
    });
    this.shared.api.GetReports().subscribe((data: { obj: { Reference: string; header: string | undefined; storeValues: any; TotalReport: any; storeorgroup: any; groups: any; FromDate: undefined; ToDate: undefined; dateType: any; dealType: any; saleType: any; financeType: any; }; }) => {
      if (data.obj.Reference == 'F&I Manager Rankings') {
        if (data.obj.header == undefined) {
          this.storeIds = data.obj.storeValues;
          this.TotalReport = data.obj.TotalReport;
          this.storeorgrp = data.obj.storeorgroup;
          this.groups = data.obj.groups
          if (data.obj.FromDate != undefined && data.obj.ToDate != undefined) {
            this.FromDate = data.obj.FromDate;
            this.ToDate = data.obj.ToDate;
            this.storeIds = data.obj.storeValues;
            this.dateType = data.obj.dateType;
            this.dealType = data.obj.dealType;
            this.saleType = data.obj.saleType;
            this.financeType = data.obj.financeType;
            this.storeorgrp == 'G' && (this.columnName != 'Rank' && this.columnName != 'fimanager' && this.columnName != 'StoreName') ? this.columnState = 'asc' : '';
            this.GetData(this.columnName, this.columnState);
          } else {
            this.FromDate = data.obj.FromDate;
            this.ToDate = data.obj.ToDate;
            this.storeIds = data.obj.storeValues;
            this.dateType = data.obj.dateType;
            this.dealType = data.obj.dealType;
            this.saleType = data.obj.saleType;
            this.financeType = data.obj.financeType;
            this.storeorgrp == 'G' && (this.columnName != 'Rank' && this.columnName != 'fimanager' && this.columnName != 'StoreName') ? this.columnState = 'asc' : '';

            this.GetData(this.columnName, this.columnState);
          }
        }
        else {
          if (data.obj.header == 'Yes') {
            this.storeIds = data.obj.storeValues;
            //console.log(this.storeIds);
            this.GetData(this.columnName, this.columnState);
          }
        }
        const headerdata = {
          title: 'F&I Manager Rankings',
          path1: '',
          path2: '',
          path3: '',
          stores: this.storeIds,
          toporbottom: this.TotalReport,
          datetype: this.dateType,
          fromdate: this.FromDate,
          todate: this.ToDate,
          groups: this.groups,
          storeorgroup: this.storeorgrp,
          dealType: this.dealType,
          saleType: this.saleType,
          financeType: this.financeType,
          otherstoreids: this.otherStoreIds
        };
        this.shared.api.SetHeaderData({
          obj: headerdata,
        });
        this.header = [{
          type: 'Bar', storeIds: this.storeIds, storeorgroup: this.storeorgrp, dealType: this.dealType,
          saleType: this.saleType,
          financeType: this.financeType, groups: this.groups,
          otherstoreids: this.otherStoreIds
        }]
      }
    });
    this.excel = this.shared.api.getExportToExcelAllReports().subscribe((res: { obj: { state: boolean; title: string; }; }) => {
      this.FIMstate = res.obj.state;
      if (res.obj.title == 'F&I Manager Rankings') {
        if (res.obj.state == true) {
          this.exportToExcel();
        }
      }
    });

    this.shared.api.getExportToEmailPDFAllReports().subscribe((res: { obj: { title: string; stateEmailPdf: boolean; }; }) => {
      if (res.obj.title == 'F&I Manager Rankings') {
        if (res.obj.stateEmailPdf == true) {
          // this.sendEmailData(res.obj.Email, res.obj.notes, res.obj.from);
        }
      }
    });
    this.shared.api.getExportToPrintAllReports().subscribe((res: { obj: { title: string; statePrint: boolean; }; }) => {
      if (res.obj.title == 'F&I Manager Rankings') {
        if (res.obj.statePrint == true) {
          // this.GetPrintData();
        }
      }
    });
    this.shared.api.getExportToPDFAllReports().subscribe((res: { obj: { title: string; statePDF: boolean; }; }) => {
      if (res.obj.title == 'F&I Manager Rankings') {
        if (res.obj.statePDF == true) {
          // this.generatePDF();
        }
      }
    });
  }

  ngOnDestroy() {
    if (this.excel != undefined) {
      this.excel.unsubscribe()
    }
  }
  displaytime() {
    if (this.DateType == 'PM') {
      return '(Same Month PY)';
    }
    else if (this.DateType == 'LM') {
      return '(Last Month)';
    }
    else if (this.DateType == 'LY') {
      return '(Last Year)';
    }
    else if (this.DateType == 'C') {
      return this.shared.datePipe.transform(this.FromDate, 'MM/dd/yyyy') + ' - ' + this.shared.datePipe.transform(this.ToDate, 'MM/dd/yyyy')
    }
    return '(' + this.DateType + ')';
  }

  StoresData(data: any) {
    this.storeIds = data.storeids;
    this.groupId = data.groupId;
    this.storename = data.storename;
    this.groupName = data.groupName;
    this.storecount = data.storecount;
    this.storedisplayname = data.storedisplayname;
    this.otherStoreIds = data.otherStoreIds;

  }



  Scrollpercent: any = 0;
  scrollCurrentposition: any = 0
  @ViewChild('scrollcent') scrollcent!: ElementRef;

  updateVerticalScroll(event: any): void {

    this.scrollCurrentposition = event.target.scrollTop
    const scrollDemo = document.querySelector('#scrollcent') as HTMLElement;
    this.Scrollpercent = Math.round(
      (event.target.scrollTop /
        (event.target.scrollHeight - scrollDemo.clientHeight)) *
      100
    );
  }

  selBlock: any;



  index = '';
  commentobj: any = {};


  // close(data: any) {
  //   // //console.log(data);
  //   this.index = '';
  // }


  //////////////////////  REPORT CODE /////////////////////////////////////////////
  // === State used in HTML ===
  Bar: boolean = true;
  activePopover: number = -1;
  neworused: string[] = ['New', 'Used'];
  DateType: string = 'MTD';
  custom: boolean = false;
  // FromDate: any;
  // ToDate: any;
  bsRangeValue!: Date[];
  minDate!: Date;
  maxDate!: Date;


  //  stores: any[] = [];
  selectedstorevalues: any[] = [];
  storeName: string = '';
  //  groupName: string = '';
  // groups: any[] = [];
  selectedGroups: any[] = [];
  AllStores: boolean = true;
  AllGroups: boolean = true;

  storeorgroup: any = ['G'];
  retailorlease: any = ['Retail', 'Lease', 'Misc', 'Special Order'];
  financetype: any = ['Finance', 'Cash', 'Lease'];

  // constructor(private pipe: DatePipe) {}

  // === HTML Interactions ===
  // @HostListener('document:click', ['$event'])
  // onDocumentClick(event: MouseEvent) {
  //   const clickedInside = (event.target as HTMLElement).closest('.dropdown-toggle, .reportstores-card, .timeframe');
  //   if (!clickedInside) {
  //     this.activePopover = -1;
  //   }
  // }
  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent) {
    const clickedInside = (event.target as HTMLElement).closest('.dropdown-toggle, .reportstores-card, .timeframe');
    if (!clickedInside) {
      this.activePopover = -1;
    }
  }
  togglePopover(popoverIndex: number) {
    this.activePopover = this.activePopover === popoverIndex ? -1 : popoverIndex;
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
    this.storesFilterData.otherStoreIds = this.otherStoreIds;
    this.storesFilterData.otherStoresArray = this.otherStoresArray;


    this.storesFilterData = {
      groupsArray: this.groupsArray,
      groupId: this.groupId,
      storesArray: this.stores,
      storeids: this.storeIds,
      groupName: this.groupName,
      storename: this.storename,
      storecount: this.storecount,
      storedisplayname: this.storedisplayname,
      'type': 'M', 'others': 'Y', otherStoresArray: this.otherStoresArray,
      otherStoreIds: this.otherStoreIds
    };

    // this.setHeaderData();
    // this.GetData();

  }



  // Deal type & status
  multipleorsingle(block: string, val: string) {
    if (block === 'NU') {

      this.toggleSelection(this.neworused, val);

    }

    if (block === 'RL') {
      this.toggleSelection(this.retailorlease, val);
    }
    if (block === 'DS') {
      this.toggleSelection(this.financetype, val);
    }
  }

  private toggleSelection(arr: any[], val: string) {
    const idx = arr.indexOf(val);
    if (idx >= 0) {
      arr.splice(idx, 1);
    } else {
      arr.push(val);
    }
  }

  storeorgroups(_block: any, val: string) {
    this.storeorgroup = [val];
  }


  initializeDates(type: any) {
    let dates: any = this.setdates.setDates(type)
    this.FromDate = dates[0];
    this.ToDate = dates[1];
    localStorage.setItem('time', type);
  }


  SetDates(type: any, block?: any) {
    if (block == 'B') {
      (<HTMLInputElement>document.getElementById('DateOfBirthBar')).click();

    }
    this.DateType = type;
    localStorage.setItem('time', this.DateType);
    let today = new Date();

    let enddate = new Date(today.setDate(today.getDate() - 1));
    let dt = new Date(today.setDate(today.getDate()));
    if (type == 'MTD') {
      this.custom = false;

      this.FromDate =
        ('0' + (enddate.getMonth() + 1)).slice(-2) +
        '-01' +
        '-' +
        enddate.getFullYear();
      this.ToDate =
        ('0' + (enddate.getMonth() + 1)).slice(-2) +
        '-' +
        ('0' + enddate.getDate()).slice(-2) +
        '-' +
        enddate.getFullYear();
      this.bsRangeValue = [this.FromDate, this.ToDate];
    }
    if (type == 'QTD') {
      this.custom = false;
      if (enddate.getMonth() == 0) {
        this.FromDate = '10-01-' + (enddate.getFullYear() - 1);
        this.ToDate = '12-31-' + (enddate.getFullYear() - 1);
        // console.log(this.FromDate, this.ToDate, 'jan Block');
        this.bsRangeValue = [this.FromDate, this.ToDate];
      } else {
        let d = new Date(enddate)
        d.setMonth(d.getMonth() - 3)
        let localstringdate = d.toISOString();
        this.FromDate = this.shared.datePipe.transform(localstringdate, 'MM-dd-yyyy')
        // localstringdate.substring(3,5) +'-' + localstringdate.substring(0,2) +'-' + localstringdate.substring(6)
        // this.FromDate =
        //   ('0' + (enddate.getMonth() - 2)).slice(-2) +
        //   '-01' +
        //   '-' +
        //   enddate.getFullYear();
        this.ToDate =
          ('0' + (enddate.getMonth() + 1)).slice(-2) +
          '-' +
          ('0' + enddate.getDate()).slice(-2) +
          '-' +
          enddate.getFullYear();
        console.log(d.toISOString(), this.FromDate, this.ToDate, 'Not jan Block');
        this.bsRangeValue = [this.FromDate, this.ToDate];
      }
    }
    if (type == 'YTD') {
      this.custom = false;

      this.FromDate =
        ('0' + 1).slice(-2) +
        '-' +
        ('0' + (enddate.getDate() - 1)).slice(-2) +
        '-' +
        enddate.getFullYear();
      this.ToDate =
        ('0' + (enddate.getMonth() + 1)).slice(-2) +
        '-' +
        ('0' + enddate.getDate()).slice(-2) +
        '-' +
        enddate.getFullYear();
      this.bsRangeValue = [this.FromDate, this.ToDate];
    }
    if (type == 'PYTD') {
      this.custom = false;
      this.FromDate = '01-01-' + (enddate.getFullYear() - 1);
      this.ToDate = ('0' + (enddate.getMonth() + 1)).slice(-2) + '-' + ('0' + enddate.getDate()).slice(-2) + '-' + (enddate.getFullYear() - 1);
      this.bsRangeValue = [new Date(this.FromDate), new Date(this.ToDate)];
    }
    if (type == 'PM') {
      this.custom = false;

      this.FromDate = ('0' + (enddate.getMonth() + 1)).slice(-2) +
        '-01' +
        '-' + (enddate.getFullYear() - 1);
      var lastDayOfMonth = new Date(
        enddate.getFullYear() - 1,
        enddate.getMonth() + 1,
        0
      );

      this.ToDate =
        ('0' + (enddate.getMonth() + 1)).slice(-2) +
        '-' +
        ('0' + lastDayOfMonth.getDate()).slice(-2) +
        '-' + (enddate.getFullYear() - 1);
      this.bsRangeValue = [this.FromDate, this.ToDate];
    }
    if (type == 'LM') {
      this.custom = false;
      if (enddate.getMonth() == 0) {
        this.FromDate = '12-01-' + (enddate.getFullYear() - 1);
        this.ToDate = '12-31-' + (enddate.getFullYear() - 1);
        this.bsRangeValue = [this.FromDate, this.ToDate];
      } else {
        this.FromDate =
          ('0' + enddate.getMonth()).slice(-2) +
          '-01' +
          '-' +
          enddate.getFullYear();
        var lastDayOfMonth = new Date(
          enddate.getFullYear(),
          enddate.getMonth(),
          0
        );
        this.ToDate =
          ('0' + enddate.getMonth()).slice(-2) +
          '-' +
          ('0' + lastDayOfMonth.getDate()).slice(-2) +
          '-' +
          enddate.getFullYear();
        this.bsRangeValue = [this.FromDate, this.ToDate];
      }
    }
    if (type == 'LY') {
      this.custom = false;
      this.FromDate = '01-01-' + (enddate.getFullYear() - 1);
      this.ToDate = '12-31-' + (enddate.getFullYear() - 1);
      this.bsRangeValue = [this.FromDate, this.ToDate];
    }
    if (type == '90') {
      this.custom = false;


      dt.setMonth(dt.getMonth() - 3);


      this.FromDate = this.shared.datePipe.transform(dt, 'MM-dd-yyyy')
      this.ToDate =
        ('0' + (enddate.getMonth() + 1)).slice(-2) +
        '-' +
        ('0' + enddate.getDate()).slice(-2) +
        '-' +
        enddate.getFullYear();
      this.bsRangeValue = [this.FromDate, this.ToDate];
      // console.log(enddate,this.ToDate,this.FromDate);

    }
    // console.log(this.FromDate);
    // console.log(this.ToDate);
  }
  dateRangeCreated($event: any) {
    if ($event) {
      this.FromDate = this.shared.datePipe.transform($event[0], 'MM-dd-yyyy');
      this.ToDate = this.shared.datePipe.transform($event[1], 'MM-dd-yyyy');
      if (this.DateType === 'C') this.custom = true;
    }
  }

  close() {
    this.shared.ngbmodal.dismissAll();
  }

  // Final Apply
  viewreport() {
    this.activePopover = -1;
    console.log('Apply clicked with:', {
      FromDate: this.FromDate,
      ToDate: this.ToDate,
      Stores: this.selectedstorevalues,
      Groups: this.selectedGroups,
      StoreOrGroup: this.storeorgroup,
      // SaleType: this.retailorlease,
      financetype: this.financetype,

      dealType: this.dealType,
      saleType: this.saleType,
      financeType: this.financeType
    });


    if (this.storeIds.length === 0 && this.otherStoreIds.lenth === 0) {

      this.toast.show('Please select any store', 'warning', 'Warning');
    }
    else if (this.retailorlease.length == 0) {

      this.toast.show('Please select any one Deal Type', 'warning', 'Warning');
    }
    else if (this.neworused.length == 0) {

      this.toast.show('Please Select Atleast One Sale Type', 'warning', 'Warning');
    }
    else if (this.financetype.length == 0) {

      this.toast.show('Please Select Atleast One Finance Type', 'warning', 'Warning');
    }
    else {
      const data = {
        Reference: 'F&I Manager Ranking',
        FromDate: this.FromDate,
        ToDate: this.ToDate,
        // TotalReport: this.toporbottom[0],
        storeValues: this.selectedstorevalues.toString(),
        // == '' ? '0': this.selectedstorevalues.toString(),
        dateType: this.DateType,
        groups: this.selectedGroups.toString(),
        storeorgroup: this.storeorgroup.toString(),
        // saleType: this.retailorlease.toString(),
        financetype: this.financetype,
        dealType: this.dealType,
        saleType: this.saleType
      };
      this.shared.api.SetReports({
        obj: data,
      });
      this.close();
      this.GetData(this.columnName, this.columnState);


    }

  }
  ExcelStoreNames: any = []
  exportToExcel(): void {
    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('F&I Manager Rankings');
    const title = worksheet.addRow(['F&I Manager Rankings']);
    title.font = { size: 14, bold: true, name: 'Arial' };
    title.alignment = { vertical: 'middle', horizontal: 'center' };
    worksheet.mergeCells('A1:J1');
    worksheet.addRow([]);
    const formattedFromDate = this.shared.datePipe.transform(this.FromDate, 'dd-MMM-yyyy');
    const formattedToDate = this.shared.datePipe.transform(this.ToDate, 'dd-MMM-yyyy');

    this.ExcelStoreNames = []
    let storeNames: any[] = [];
    // let store: any = [];
    // this.storeIds && this.storeIds.toString().indexOf(',') > 0 ? store = this.storeIds.toString().split(',') : store=this.storeIds
    // console.log(this.storeIds, this.groups, store, '...............................');
    //     storeNames = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groups)[0].Stores.filter((item: any) => store.includes(item.ID.toString()));

    //     console.log(storeNames, 'storeNames');
    //     if (store.length == this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groups)[0].Stores.length)
    //     {
    //     //    { this.ExcelStoreNames = 'All Stores' }
    //     // else 
    //     //   {
    //          this.ExcelStoreNames = storeNames.map(function (a: any) { return a.storename; }); 
    //         }

    //  console.log( this.ExcelStoreNames , 'exclstoreNames');

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
    console.log(this.ExcelStoreNames, 'exclstoreNames');
    const filters = [
      { name: 'Stores:', values: storeValue },
      { name: 'Time Frame:', values: `${formattedFromDate} to ${formattedToDate}` },
      { name: 'Rank By:', values: this.storeorgroup == 'S' ? 'Store' : 'Group' },
      { name: 'New/Used:', values: this.neworused || 'All' },
      { name: 'Deal Type:', values: this.retailorlease || 'All' },
      { name: 'Finance Type:', values: this.financetype || 'All' },
    ];
    let currentRow = worksheet.lastRow?.number ?? worksheet.rowCount;
    filters.forEach((filter) => {
      currentRow++;
      let value = Array.isArray(filter.values)
        ? filter.values.join(', ')
        : filter.values;
      const row = worksheet.addRow([filter.name, value]);
      row.getCell(1).font = { bold: true, name: 'Arial', size: 10 };
      row.getCell(2).font = { name: 'Arial', size: 10, color: { argb: 'FF1F497D' } }; // blue color for values
      worksheet.mergeCells(`B${currentRow}:F${currentRow}`);
    });
    worksheet.addRow([]);
    const firstHeader = ['', '', '', 'Back Gross', '', '', 'Unit Count', '', ''];
    const headerRow1 = worksheet.addRow(firstHeader);
    const secondHeader = [
      'Rank', 'Finance Manager', 'Store Name',
      'PVR', 'Gross', 'Pace',
      'New', 'Used', 'Total', 'Pace'
    ];
    const headerRow2 = worksheet.addRow(secondHeader);
    const headerRow1Index = headerRow1.number;
    worksheet.mergeCells(`A${headerRow1Index}`);
    worksheet.mergeCells(`B${headerRow1Index}`);
    worksheet.mergeCells(`C${headerRow1Index}`);
    worksheet.mergeCells(`D${headerRow1Index}:F${headerRow1Index}`); // Back Gross
    worksheet.mergeCells(`G${headerRow1Index}:J${headerRow1Index}`); // Unit Count
    [headerRow1, headerRow2].forEach(r => {
      r.height = 22;
      r.eachCell({ includeEmpty: false }, (cell) => {
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FF2F5597' }, // deep blue
        };
      });
    });
    const bindingHeaders = [
      'Rank', 'fimanager', 'StoreName',
      'BackgrossPVR', 'BackGross',
      'BackGrossPace', 'New', 'Used', 'Total', 'Pace'
    ];
    const currencyFields = ['BackgrossPVR', 'BackGross',
      'BackGrossPace'];
    this.FIManagerData.forEach((info: any) => {
      const rowData = bindingHeaders.map((key) => {
        const val = info[key];
        return (val === 0 || val == null || val === '') ? '-' : val;
      });
      const dataRow = worksheet.addRow(rowData);
      bindingHeaders.forEach((key, index) => {
        const cell = dataRow.getCell(index + 1);

        if (currencyFields.includes(key) && typeof cell.value === 'number') {
          cell.numFmt = '"$"#,##0.00';
          cell.alignment = { horizontal: 'right', vertical: 'middle' };
        } else {
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
        }
      });
    });
    worksheet.columns.forEach(col => col.width = 25);
    workbook.xlsx.writeBuffer().then(buffer => {
      this.shared.exportToExcel(workbook, 'F&I Manager Rankings');
    });
  }


}

