import { Component, OnInit, ViewChild, ElementRef, HostListener } from '@angular/core';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { Setdates } from '../../../../Core/Providers/SetDates/setdates';
import { BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { common } from '../../../../common';
import { Stores } from '../../../../CommonFilters/stores/stores';
import { DateRangePicker } from '../../../../CommonFilters/date-range-picker/date-range-picker';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
import { NgbModal } from '@ng-bootstrap/ng-bootstrap';
// import { SalesgrossDetailsComponent } from '../../Sales/Gross/salesgross-details/salesgross-details.component';
const EXCEL_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
const EXCEL_EXTENSION = '.xlsx';

@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [SharedModule, BsDatepickerModule, DateRangePicker, Stores],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
})
export class Dashboard {
  IndividualSalesPersonsData: any = [];
  FromDate: any = '';
  ToDate: any = '';
  DupFromDate: any = '';
  DupToDate: any = ''
  minDate!: Date;
  maxDate!: Date;
  DateType: any = 'MTD';
  displaytime: any = '';
  NoData: boolean = false;
  //  storeIds: any = []
  stores: any = []
  groupsArray: any = [];
  storename: any = ''
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  groupId: any = 0;
  		  otherStoresArray: any = [];
  otherStoreIds: any = [];

  storesFilterData: any = {
    'groupsArray': this.groupsArray, 'groupId': this.groupId, 'storesArray': this.stores, 'storeids': '1', 'type': 'M', 'others': 'Y',
    'groupName': this.groupName, 'storename': this.storename, storecount: null, 'storedisplayname': this.storedisplayname,
    otherStoresArray: this.otherStoresArray, otherStoreIds: this.otherStoreIds

  };

  Dates: any = {
    'FromDate': this.FromDate, 'ToDate': this.ToDate, "MaxDate": this.maxDate, 'MinDate': this.minDate, 'DateType': this.DateType, 'DisplayTime': this.displaytime,
    Types: [
      { 'code': 'MTD', 'name': 'MTD' },
      { 'code': 'QTD', 'name': 'QTD' },
      { 'code': 'YTD', 'name': 'YTD' },
      { 'code': 'PYTD', 'name': 'PYTD' },
      { 'code': 'LY', 'name': 'Last Year' },
      { 'code': 'LM', 'name': 'Last Month' },
      { 'code': 'PM', 'name': 'Same Month PY' },
    ]
  }

  // filters
  TotalReport: any = 'T';
  storeIds!: any;
  dateType: any = 'MTD';
  DupDateType: any = 'MTD';

  groups: any = 0;
  storeorgrp: any = 'S';
  saleType: any = 'Retail,Lease';
  retailorlease: any = this.saleType.split(',');
  columnName: any = 'Rank';
  columnState: any = 'asc';
  dealStatus: any = ['Booked', 'Finalized', 'Delivered'];

  CompleteComponentState: boolean = true;
  solutionurl: any;
  LogCount = 1;
  header: any = [
    {
      type: 'Bar',
      storeIds: this.storeIds,
      storeorgroup: this.storeorgrp,
      dealStatus: this.dealStatus,
      saleType: this.saleType,
      groups: this.groups,
    },
  ];
  popup: any = [{ type: 'Popup' }];

  constructor(public shared: Sharedservice, public setdates: Setdates, private comm: common, private toast: ToastService,private ngbmodal : NgbModal
  ) {
    this.solutionurl = this.shared.api;

    // this.initializeDates();
    this.shared.setTitle('Salesperson Rankings');
  

    if (typeof window !== 'undefined') {
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
        this.otherStoreIds = JSON.parse(localStorage.getItem('otherstoreids')!);

        }
      }
      if (this.shared.common.groupsandstores.length > 0) {
        this.groupsArray = this.shared.common.groupsandstores.filter((val: any) => val.sg_id != this.shared.common.reconID);
          this.otherStoresArray = this.shared.common.OtherStoresData[0].Stores

        this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
        this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_Name : this.groupName = ''
        this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
        // // console.log(this.stores, this.groupsArray, 'Stores and Groups');
        this.getStoresandGroupsValues()
        // this.StoresData(this.ngChanges)
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
      localStorage.setItem('stime', 'MTD')


      this.shared.setTitle(this.shared.common.titleName + '-Salesperson Rankings');
      // if (localStorage.getItem('Fav') != 'Y') {
      const data = {
        title: 'Salesperson Rankings',
        stores: this.storeIds,
        toporbottom: this.TotalReport,
        datetype: this.DateType,
        fromdate: this.FromDate,
        todate: this.ToDate,
        groups: this.groups,
        storeorgroup: this.storeorgrp,
        saleType: this.saleType.toString(),
        dealStatus: this.dealStatus,
        count: 0,
      };
      this.shared.api.SetHeaderData({
        obj: data,
      });
      this.header = [
        {
          type: 'Bar',
          storeIds: this.storeIds,
          dealStatus: this.dealStatus,
          storeorgroup: this.storeorgrp,
          fromdate: this.FromDate,
          todate: this.ToDate,
          saleType: this.saleType,
          groups: this.groups,
        },
      ];
      this.GetData('Rank', 'asc');
      this.setDates(this.DateType)

      // }
    }
  }

  ngOnInit(): void { }

  initializeDates(type: any) {
    let dates: any = this.setdates.setDates(type)
    this.FromDate = dates[0];
    this.ToDate = dates[1];
    localStorage.setItem('time', type);
  }

  // private initializeDates() {
  //   const today = new Date();
  //   const enddate = new Date(today.setDate(today.getDate() - 1));

  //   this.FromDate = `${('0' + (enddate.getMonth() + 1)).slice(-2)}/01/${enddate.getFullYear()}`;
  //   this.ToDate = `${('0' + (enddate.getMonth() + 1)).slice(-2)}/${('0' + enddate.getDate()).slice(
  //     -2
  //   )}/${enddate.getFullYear()}`;
  // }
  // need this function
  openDetails(data: any) {
    //     const DetailsSalesPeron = this.shared.ngbmodal.open(
    //       SalesgrossDetailsComponent,
    //       {
    //         size: 'xxl',
    //         backdrop: 'static',
    //       }
    //     );
    //     DetailsSalesPeron.componentInstance.Salesdetails = [
    //       {
    //         StartDate: this.FromDate,
    //         EndDate: this.ToDate,
    //         dealtype:[
    //           "New",
    //           "Used"
    //       ],
    //         saletype: this.saleType,
    //         dealstatus: [
    //           "Delivered",
    //           "Booked",
    //           "Finalized"
    //       ],
    //         var1: 'store',
    //         var2: 'salesperson',
    //         var3: '',
    //         var1Value: data.StoreName,
    //         var2Value: data.SPTrimName,
    //         var3Value: '',
    //         userName: data.SPTrimName,
    //         FinanceManager: "0",
    // SalesManager: "0",
    // SalesPerson: "0"
    //       },
    //     ];
    //     DetailsSalesPeron.result.then(
    //       (data:any) => { },
    //       (reason:any) => {
    //         // on dismiss
    //       }
    //     );
  }
  Favreports: any = [];



  StoresData(data: any) {

    console.log(data, 'Data');

    this.storeIds = data.storeids;
    this.groupId = data.groupId;
    this.storename = data.storename;
    this.groupName = data.groupName;
    this.storecount = data.storecount;
    this.storedisplayname = data.storedisplayname;
	      this.otherStoreIds = data.otherStoreIds;

  }

  // tabClick(col_Name: any, Col_state: any) {
  //   if (this.columnName == col_Name) {
  //     if (Col_state == 'asc') {
  //       this.columnState = 'desc';
  //       this.GetData(this.columnName, this.columnState);
  //     } else {
  //       this.columnState = 'asc';
  //       this.GetData(this.columnName, this.columnState);
  //     }
  //   } else {
  //     if (
  //       this.storeorgrp == 'G' &&
  //       col_Name != 'Rank' &&
  //       col_Name != 'ServiceAdvisor' &&
  //       col_Name != 'StoreName'
  //     ) {
  //       this.columnState = 'desc';
  //       this.columnName = col_Name;
  //       this.GetData(this.columnName, this.columnState);
  //     } else {
  //       this.columnState = 'desc';
  //       this.columnName = col_Name;
  //       this.GetData(this.columnName, this.columnState);
  //     }
  //   }
  // }
  tabClick(col_Name: any) {

    // First click on a column
    if (this.columnName !== col_Name) {
      this.columnName = col_Name;
      this.columnState = 'asc';
    }

    else {
      this.columnState = this.columnState === 'asc' ? 'desc' : 'asc';
    }

    this.GetData(this.columnName, this.columnState);
  }
  GetData(sortdata?: any, sortstate?: any) {
    console.log(sortdata, sortstate, this.storeIds);
  this.DupFromDate = this.FromDate;
    this.DupToDate = this.ToDate
    this.DupDateType= this.DateType

    this.IndividualSalesPersonsData = [];
    this.shared.spinner.show();
    const obj = {
      UserID: 0,
      StartDate: this.FromDate,
      EndDate: this.ToDate,
      StoreID: [...this.storeIds, ...this.otherStoreIds],
      Exp: sortdata,
      OrderType: sortstate,
      RankBy: this.storeorgrp,
      DealType: this.saleType,
      DealStatus: this.dealStatus.toString(),
    };
    let startFrom = new Date().getTime();
    const curl = this.shared.getEnviUrl() + 'GetSalesPersonsRankings';
    this.shared.api
      .postmethod(this.shared.common.routeEndpoint + 'GetSalesPersonsRankings', obj)
      .subscribe(
        (res) => {
          // const currentTitle = document.title;
          // this.shared.api.logSaving(curl, {}, '', res.message, currentTitle);
          if (res.status == 200) {
            if (res.response != undefined) {
              if (res.response.length > 0) {
                this.IndividualSalesPersonsData = res.response;

                this.shared.spinner.hide();

                this.NoData = false;
              } else {
                this.shared.spinner.hide();
                this.NoData = true;
              }
            } else {
              //  this.shared.toaster.error('Empty Response');
              this.shared.spinner.hide();
              this.NoData = true;

              // this.shared.spinnerLoader=false;
            }
          } else {
            //  this.shared.toaster.error(res.status);
            this.shared.spinner.hide();
            this.NoData = true;
            // this.shared.spinnerLoader=false;
          }
        },
        (error) => {
          //  this.shared.toaster.error('502 Bad Gate Way Error', '');
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

  isDesc: boolean = false;
  column: string = 'CategoryName';
  sort(property: any) {
    this.isDesc = !this.isDesc; //change the direction
    this.column = property;
    let direction = this.isDesc ? 1 : -1;
    if (this.TotalReport == 'T') {
      var arr = this.IndividualSalesPersonsData.slice(1, this.IndividualSalesPersonsData.length);
      arr.sort(function (a: any, b: any) {
        if (a[property] < b[property]) {
          return -1 * direction;
        } else if (a[property] > b[property]) {
          return 1 * direction;
        } else {
          return 0;
        }
      });
      arr.unshift(this.IndividualSalesPersonsData[0]);
      this.IndividualSalesPersonsData = arr;
    } else {
      var arr = this.IndividualSalesPersonsData.slice(0, -1);
      arr.sort(function (a: any, b: any) {
        if (a[property] < b[property]) {
          return -1 * direction;
        } else if (a[property] > b[property]) {
          return 1 * direction;
        } else {
          return 0;
        }
      });
      arr.push(this.IndividualSalesPersonsData[this.IndividualSalesPersonsData.length - 1]);
      this.IndividualSalesPersonsData = arr;
    }
  }
  SPRstate: any;
  ngAfterViewInit(): void {
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.comm.pageName == 'Salesperson Rankings') {
        if (res.obj.storesData != undefined) {
          this.groupsArray = res.obj.storesData;
          this.otherStoresArray = this.shared.common.OtherStoresData[0].Stores

          // this.groupId = this.ngChanges.groups;
          this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
          // this.storeIds = this.ngChanges.storeIds;
          this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_name : this.groupName = ''
          this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
          // // console.log(this.stores, this.groupsArray, 'Stores and Groups');
          this.getStoresandGroupsValues()
        }
      }
    })
    this.shared.api.GetReportOpening().subscribe((res) => {
      // // console.log(res);
      if (res.obj.Module == 'Salesperson Rankings') {
        document.getElementById('report')?.click();
      }
    });
    this.shared.api.GetReports().subscribe((data) => {
      if (data.obj.Reference == 'Salesperson Rankings') {
        if (data.obj.header == undefined) {
          this.TotalReport = data.obj.TotalReport;
          this.storeIds = data.obj.storeValues;

          this.dealStatus = data.obj.dealStatus;
          if (this.storeorgrp != data.obj.storeorgroup) {
            this.columnName = 'Rank';
            // this.columnState = 'asc';
            this.storeorgrp = data.obj.storeorgroup;
          } else {
            this.storeorgrp = data.obj.storeorgroup;
            // this.columnState = 'desc';
          }
          this.storeorgrp = data.obj.storeorgroup;
          this.saleType = data.obj.saleType;
          this.groups = data.obj.groups;
          if (data.obj.FromDate != undefined && data.obj.ToDate != undefined) {
            this.FromDate = data.obj.FromDate;
            this.ToDate = data.obj.ToDate;
            this.storeIds = data.obj.storeValues;
            this.dateType = data.obj.dateType;
            this.GetData(this.columnName, this.columnState);
          } else {
            this.FromDate = data.obj.FromDate;
            this.ToDate = data.obj.ToDate;
            this.storeIds = data.obj.storeValues;
            this.dateType = data.obj.dateType;
            this.GetData(this.columnName, this.columnState);
          }
        } else {
          if (data.obj.header == 'Yes') {
            this.storeIds = '1';
            // console.log(this.storeIds);
            this.GetData(this.columnName, this.columnState);
          }
        }
        const headerdata = {
          title: 'Salesperson Rankings',

          stores: this.storeIds,
          toporbottom: this.TotalReport,
          datetype: this.dateType,
          fromdate: this.FromDate,
          todate: this.ToDate,
          groups: this.groups,
          storeorgroup: this.storeorgrp,
          saleType: this.saleType,
          dealStatus: this.dealStatus,
          otherstoreids: this.otherStoreIds
        };
        this.shared.api.SetHeaderData({
          obj: headerdata,
        });
        this.header = [
          {
            type: 'Bar',
            storeIds: this.storeIds,
            storeorgroup: this.storeorgrp,
            fromdate: this.FromDate,
            todate: this.ToDate,
            saleType: this.saleType,
            groups: this.groups,
            dealStatus: this.dealStatus,
            otherstoreids: this.otherStoreIds
          },
        ];
      }
    });
    this.shared.api.getExportToExcelAllReports().subscribe((res) => {
      this.SPRstate = res.obj.state;
      if (res.obj.title == 'Salesperson Rankings') {
        if (res.obj.state == true) {
          this.exportToExcel();
        }
      }
    });

    this.shared.api.getExportToPrintAllReports().subscribe((res) => {
      if (res.obj.title == 'Salesperson Rankings') {
        if (res.obj.statePrint == true) {
          // this.GetPrintData();
        }
      }
    });

    this.shared.api.getExportToPDFAllReports().subscribe((res) => {
      if (res.obj.title == 'Salesperson Rankings') {
        if (res.obj.statePDF == true) {
          // this.generatePDF();
        }
      }
    });
    this.shared.api.getExportToEmailPDFAllReports().subscribe((res) => {
      if (res.obj.title == 'Salesperson Rankings') {
        if (res.obj.stateEmailPdf == true) {
          // this.sendEmailData(res.obj.Email, res.obj.notes, res.obj.from);
        }
      }
    });
  }


  getStoresandGroupsValues() {


    //   this.stores = this.comm.groupsandstores.filter((v: any) => v.sg_id == this.groupId[0])[0].Stores;
    //   // console.log( this.stores)
    //   this.storeIds = []
    //   this.storeIds.push(1)
    //  // console.log( this.storeIds.length)
    //   // let data = this.comm.completeUserDetails;
    //   // data.Store_Ids.indexOf(',') > 0 ? this.storeIds.push(parseInt(data.Store_Ids.split(',')[0])) : this.storeIds.push(data.Store_Ids)
    //   this.storecount = this.storeIds.length
    //   if (this.storeIds.length == 1) {
    //     this.storename = this.stores.filter((val: any) => val.ID == this.storeIds.toString())[0].storename;
    //     this.storecount = null;
    //     this.storedisplayname = this.storename
    //   }
    //   else if (this.storeIds.length == this.stores.length) {
    //     this.groupName = this.groupsArray.filter((val: any) => val.sg_id == this.groupId[0])[0].sg_name;
    //     this.storecount = null;
    //     this.storedisplayname = this.groupName;
    //   }
    //   else if (this.storeIds.length > 1) {
    //     this.storecount = this.storeIds.length;
    //     this.storedisplayname = 'Selected'
    //   }
    //   else {
    //     this.storedisplayname = 'Select'
    //   }

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

  reportOpen(temp: any) {
    // this.ngbmodalActive = this.ngbmodal.open(temp, {
    //   size: 'xl',
    //   backdrop: 'static',
    // });
  }

  openCustomPicker(datepicker: any) {
    this.DateType = 'C';
    localStorage.setItem('time', this.DateType);
    datepicker.toggle();
  }


  dateRangeCreated($event: any) {
    if ($event) {
      this.FromDate = this.shared.datePipe.transform($event[0], 'MM-dd-yyyy');
      this.ToDate = this.shared.datePipe.transform($event[1], 'MM-dd-yyyy');
      this.bsRangeValue = [this.FromDate, this.ToDate];
      if (this.DateType === 'C') this.custom = true;
    }
  }

  updatedDates(data: any) {
    // console.log(data);
    this.FromDate = data.FromDate;
    this.ToDate = data.ToDate;
    this.DateType = data.DateType;
    this.displaytime = data.DisplayTime
  }




  // comments code

  Scrollpercent: any = 0;
  scrollCurrentposition: any = 0;
  @ViewChild('scrollcent') scrollcent!: ElementRef;
  updateVerticalScroll(event: any): void {
    this.scrollCurrentposition = event.target.scrollTop;
    const scrollDemo = document.querySelector('#scrollcent') as HTMLElement;
    this.Scrollpercent = Math.round(
      (event.target.scrollTop / (event.target.scrollHeight - scrollDemo.clientHeight)) * 100
    );
  }






  //////////////////////  REPORT CODE /////////////////////////////////////////////
  // === State used in HTML ===
  Bar: boolean = true;
  activePopover: number = -1;

  custom: boolean = false;
  // FromDate: any;
  // ToDate: any;
  bsRangeValue!: Date[];



  // stores: any[] = [];
  selectedstorevalues: any[] = [];
  storeName: string = '';
  // groupName: string = '';
  // groups: any[] = [];
  selectedGroups: any[] = [];
  AllStores: boolean = true;
  AllGroups: boolean = true;

  storeorgroup: any = ['S'];
  // retailorlease: any = [];
  // saleType: any = 'Retail,Lease,Wholesale,Misc,Fleet,Demo,Special Order,Rental,Dealer Trade';

  // dealStatus: any = ['Booked', 'Finalized', 'Delivered'];

  // constructor(private pipe: DatePipe) {}

  // === HTML Interactions ===
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

  // Store selection
  allstores() {
    this.AllStores = !this.AllStores;
    this.selectedstorevalues = this.AllStores ? this.stores.map((s: { ID: any; }) => s.ID) : [];
  }


  individualStores(e: any) {

    const index = this.selectedstorevalues.findIndex((i: any) => i == e.ID);
    if (index >= 0) {
      this.selectedstorevalues.splice(index, 1);
      this.AllStores = false;
      if (this.selectedstorevalues.length == 1) {
        this.storeorgroup = 'S'
      }
      else {
        this.storeorgroup = 'G'
      }

    } else {
      this.selectedstorevalues.push(e.ID);
      if (this.selectedstorevalues.length == 1) {
        this.storeorgroup = 'S'
      }
      else {
        this.storeorgroup = 'G'
      }
      if (this.selectedstorevalues.length == this.stores.length) {
        this.AllStores = true;
      } else {
        this.AllStores = false;
      }
    }
    if (this.selectedstorevalues.length == 1) {
      this.storeName = this.stores.filter((val: any) => val.ID == this.selectedstorevalues.toString())[0].storename
    }
  }
  // Group selection
  allgroups() {
    this.AllGroups = !this.AllGroups;
    this.selectedGroups = this.AllGroups ? this.groups.map((g: { sg_id: any; }) => g.sg_id) : [];
  }

  individualgroups(g: any) {
    this.selectedGroups = [g.sg_id];
    this.groupName = g.sg_name;
    this.storeorgroup = g.sg_id === 1 ? 'S' : 'G';
  }

  // Deal type & status
  multipleorsingle(block: string, val: string) {
    if (block === 'RL') {
      this.toggleSelection(this.retailorlease, val);
    }
    if (block === 'DS') {
      this.toggleSelection(this.dealStatus, val);
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
  selBlock: any;
  screenheight: any = 0; divheight: any = 0; trposition: any = 0;
  commentopen(item: any, i: any, slblock: any = '') {
    this.screenheight = window.screen.height;
    this.divheight = (<HTMLInputElement>document.getElementById('scrollcent')).offsetHeight;
    this.trposition = (<HTMLInputElement>document.getElementById('DV_' + i)).offsetTop;
    this.index = '';
    console.log('Selected Obj :', item);
    //return
    this.selBlock = slblock + i.toString();
    this.index = i.toString();
    this.commentobj = {
      TYPE: item.SalesPerson,
      NAME: item.LABLE1,
      STORES: item.StoreName,
      STORENAME: item.StoreName,
      Month: '',
      ModuleId: '68',
      ModuleRef: 'SPR',
      state: 1,
      indexval: i,
    };
  }
  index = '';
  commentobj: any = {};
  addcmt(data: any) {
    if (data == 'A') {
      this.index = '';
      const DetailsSF = this.ngbmodal.open({
        size: 'xl',
        backdrop: 'static',
      });
      // myObject['skillItem2'] = 15;
      this.commentobj['state'] = 0;
      (DetailsSF.componentInstance.SFComments = this.commentobj),
        DetailsSF.result.then(
          (data) => {
            // console.log(data);
          },
          (reason) => {
            // console.log(reason);
            // 
            if (reason == 'O') {
              this.commentobj['state'] = 1;
              this.index = this.commentobj['indexval'];
            } else {
              this.commentobj['state'] = 1;
              this.index = this.commentobj['indexval'];
              this.GetData(this.columnName, this.columnState);


            }
            // // on dismiss

            // const Data = {
            //   state: true,
            // };
            // this.apiSrvc.setBackgroundstate({ obj: Data });
            // this.GetData();
          }
        );
    }
    if (data == 'AD') {
      this.GetData(this.columnName, this.columnState);

      // if (this.Filter == 'VariableTrendsvsBudget') {
      //   this.GetData();
      // }
      // if (this.Filter == 'VariableTrendsvsStores') {
      //   this.GetData();
      // }
    }
  }

  // Date selection
  // SetDates(type: string, block?: string) {
  //   this.DateType = type;
  //   let today = new Date();
  //   let enddate = new Date(today.setDate(today.getDate() - 1));

  //   if (type === 'MTD') {
  //     this.custom = false;
  //     this.FromDate = this.shared.datePipe.transform(new Date(enddate.getFullYear(), enddate.getMonth(), 1), 'MM-dd-yyyy');
  //     this.ToDate = this.shared.datePipe.transform(enddate, 'MM-dd-yyyy');
  //     this.bsRangeValue = [this.FromDate, this.ToDate];
  //   }
  //   if (type === 'C' && block === 'B') {
  //     // just trigger custom
  //     this.custom = true;
  //   }
  // }

  // getStoresandGroupsValues() {


  //   this.stores = this.comm.groupsandstores.filter((v: any) => v.sg_id == this.groupId[0])[0].Stores;
  //   // console.log( this.stores)
  //   this.storeIds = []
  //   this.storeIds.push(1)
  //  // console.log( this.storeIds.length)
  //   // let data = this.comm.completeUserDetails;
  //   // data.Store_Ids.indexOf(',') > 0 ? this.storeIds.push(parseInt(data.Store_Ids.split(',')[0])) : this.storeIds.push(data.Store_Ids)
  //   this.storecount = this.storeIds.length
  //   if (this.storeIds.length == 1) {
  //     this.storename = this.stores.filter((val: any) => val.ID == this.storeIds.toString())[0].storename;
  //     this.storecount = null;
  //     this.storedisplayname = this.storename
  //   }
  //   else if (this.storeIds.length == this.stores.length) {
  //     this.groupName = this.groupsArray.filter((val: any) => val.sg_id == this.groupId[0])[0].sg_name;
  //     this.storecount = null;
  //     this.storedisplayname = this.groupName;
  //   }
  //   else if (this.storeIds.length > 1) {
  //     this.storecount = this.storeIds.length;
  //     this.storedisplayname = 'Selected'
  //   }
  //   else {
  //     this.storedisplayname = 'Select'
  //   }

  //   this.storesFilterData.groupsArray = this.groupsArray;
  //   this.storesFilterData.groupId = this.groupId;
  //   this.storesFilterData.storesArray = this.stores;
  //   this.storesFilterData.storeids = this.storeIds;
  //   this.storesFilterData.groupName = this.groupName;
  //   this.storesFilterData.storename = this.storename;
  //   this.storesFilterData.storecount = this.storecount;
  //   this.storesFilterData.storedisplayname = this.storedisplayname;
  //   // this.setHeaderData();
  //   this.GetData();

  // }



  setDates(type: any) {
    // localStorage.setItem('time', type);
    // this.datevaluetype=
    // console.log(type);

    this.displaytime = '(' + this.Dates.Types.filter((val: any) => val.code == type)[0].name + ')';
    this.maxDate = new Date();
    this.minDate = new Date();
    this.minDate.setFullYear(this.maxDate.getFullYear() - 3);
    this.maxDate.setDate(this.maxDate.getDate());
    this.Dates.FromDate = this.FromDate;
    this.Dates.ToDate = this.ToDate;
    this.Dates.MinDate = this.minDate;
    this.Dates.MaxDate = this.maxDate;
    this.Dates.DateType = this.DateType;
    this.Dates.DisplayTime = this.displaytime;
  }
  // dateRangeCreated($event: any) {
  //   if ($event) {
  //     this.FromDate = this.shared.datePipe.transform($event[0], 'MM-dd-yyyy');
  //     this.ToDate = this.shared.datePipe.transform($event[1], 'MM-dd-yyyy');
  //     if (this.DateType === 'C') this.custom = true;
  //   }
  // }

  close() {
    this.shared.ngbmodal.dismissAll();
  }

  // Final Apply
  viewreport() {
    this.activePopover = -1;
    // console.log('Apply clicked with:', {
    //   FromDate: this.FromDate,
    //     ToDate: this.ToDate,
    //       Stores: this.selectedstorevalues,
    //         Groups: this.selectedGroups,
    //           StoreOrGroup: this.storeorgroup,
    //             SaleType: this.retailorlease,
    //               DealStatus: this.dealStatus
    // });

    if ((!this.storeIds || this.storeIds.length === 0) && (this.otherStoreIds.length ===0)) {

      this.toast.show('Please Select Atleast One Store', 'warning', 'Warning');
      return;
    }
    else if (this.retailorlease.length == 0) {

      this.toast.show('Please select any one Deal Type', 'warning', 'Warning');
    }
    else if (this.dealStatus.length == 0) {

      this.toast.show('Please Select Atleast One Deal Status', 'warning', 'Warning');
    }
    else {
      const data = {
        Reference: 'Salesperson Rankings',
        FromDate: this.FromDate,
        ToDate: this.ToDate,
        // TotalReport: this.toporbottom[0],
        storeValues: this.storeIds.toString(),
        // == '' ? '0': this.selectedstorevalues.toString(),
        dateType: this.DateType,
        groups: this.selectedGroups.toString(),
        storeorgroup: this.storeorgroup.toString(),
        saleType: this.retailorlease.toString(),
        dealStatus: this.dealStatus,
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
    const worksheet = workbook.addWorksheet('Salesperson Rankings');
    const title = worksheet.addRow(['Salesperson Rankings']);
    title.font = { size: 14, bold: true, name: 'Arial' };
    title.alignment = { vertical: 'middle', horizontal: 'center' };
    worksheet.mergeCells('A1:L1');
    worksheet.addRow([]);
    const formattedFromDate = this.shared.datePipe.transform(this.FromDate, 'dd-MMM-yyyy');
    const formattedToDate = this.shared.datePipe.transform(this.ToDate, 'dd-MMM-yyyy');

    // const rankByValue =
    // typeof this.storeorgroup === 'string'
    //   ? this.storeorgroup.toUpperCase() === 'S'
    //     ? 'Store'
    //     : this.storeorgroup.toUpperCase() === 'G'
    //       ? 'Group'
    //       : this.storeorgroup
    //   : 'Store';

    this.ExcelStoreNames = []
    let storeNames: any[] = [];

    //  }

    worksheet.addRow([]);
    // let storeValue = 'All Stores';
    // if (
    //   this.storeIds &&
    //   this.storeIds.length > 0 &&
    //   this.storeIds.length !== this.stores.length
    // ) {
    //   storeValue = this.stores
    //     .filter((s: any) => this.storeIds.includes(s.ID))
    //     .map((s: any) => s.storename)
    //     .join(', ');
    // }
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
      { name: 'Time Frame:', values: `${formattedFromDate} to ${formattedToDate}` },
      { name: 'Rank By:', values: this.storeorgroup == 'S' ? 'Store' : 'Group' },
      // { name: 'New/Used:', values: this.neworused || 'All' },
      { name: 'Deal Type:', values: this.retailorlease || 'All' },
      { name: 'Deal Status:', values: this.dealStatus.toString().replace('Finalized', 'Finalized') || 'All' },
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



    const firstHeader = ['', '', '', 'Unit Count', '', '', '', '', 'Total Gross', '', '', ''];
    const headerRow1 = worksheet.addRow(firstHeader);


    const secondHeader = [
      'Rank', 'Salesperson', 'Store Name',

      'New', 'Used', 'Total', 'Pace', '90 Day Avg', 'Total', 'Pace', 'PVR', '90 Day Avg'
    ];
    const headerRow2 = worksheet.addRow(secondHeader);

    const headerRow1Index = headerRow1.number;
    const headerRow2Index = headerRow2.number;


    worksheet.mergeCells(`A${headerRow1Index}`);
    worksheet.mergeCells(`B${headerRow1Index}`);
    worksheet.mergeCells(`C${headerRow1Index}`);
    worksheet.mergeCells(`D${headerRow1Index}:H${headerRow1Index}`); // Back Gross
    worksheet.mergeCells(`I${headerRow1Index}:L${headerRow1Index}`); // Unit Count

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
      'Rank', 'SalesPerson', 'StoreName',
      'MTD_NEW', 'MTD_USED',
      'MTD_Total', 'Pace', 'UnitDayAvg', 'Gross', 'GrossPace', 'PVR', 'GrossDayAvg'
    ];

    const currencyFields = ['Gross', 'GrossPace', 'PVR', 'GrossDayAvg'];

    this.IndividualSalesPersonsData.forEach((info: any) => {
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
      this.shared.exportToExcel(workbook, 'Salesperson Rankings');
    });
  }
}


