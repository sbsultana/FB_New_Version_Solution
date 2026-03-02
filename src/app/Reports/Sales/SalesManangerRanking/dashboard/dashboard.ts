import { Component, OnInit, ViewChild, ElementRef, HostListener } from '@angular/core';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { Setdates } from '../../../../Core/Providers/SetDates/setdates';
import { BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { common } from '../../../../common';
 import { Stores } from '../../../../CommonFilters/stores/stores';
import { DateRangePicker } from '../../../../CommonFilters/date-range-picker/date-range-picker';
import { Subscription } from 'rxjs';

import { Workbook } from 'exceljs';
import * as FileSaver from 'file-saver';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
const EXCEL_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
const EXCEL_EXTENSION = '.xlsx';
@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [SharedModule, BsDatepickerModule,Stores,DateRangePicker],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
})
export class Dashboard {
  SalesPersonsData: any = [];
  FIManagerData: any = [];
  TotalSalesPersonsData: any = [];
  // FromDate: any;
  // ToDate: any;
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
  groups: any = 1;

  header: any = [{ type: 'Bar', storeIds: this.storeIds,storeorgroup: this.storeorgrp, dealType: this.dealType,
  saleType: this.saleType,
  financeType: this.financeType ,groups:this.groups}]
  popup: any = [{ type: 'Popup' }];

  // solutionurl: any = environment.apiUrl;
  LogCount = 1;
  StoreVal: any = '0';


 
  stores: any = [];
  columnName: any = 'Rank';
  columnState: any = 'asc';
  groupsArray: any = [];
  storename: any = ''
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  groupId: any = 8;

  storesFilterData: any = {
    'groupsArray': this.groupsArray, 'groupId': this.groupId, 'storesArray': this.stores, 'storeids': '1', 'type': 'M', 'others': 'N',
    'groupName': this.groupName, 'storename': this.storename, storecount: null, 'storedisplayname': this.storedisplayname
  };
  FromDate: any = '';
  ToDate: any = '';
  minDate!: Date;
  maxDate!: Date;
  DateType: any = 'MTD';
  displaytime: any = '';


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
  constructor(
 public shared: Sharedservice, public setdates: Setdates,private toast: ToastService,
  ) {


    if (localStorage.getItem('flag') == 'V') {
        this.storeIds = [];
        console.log(JSON.parse(localStorage.getItem('userInfo')!), JSON.parse(localStorage.getItem('userInfo')!).user_Info, 'Widget Stores............');
        this.groupId = JSON.parse(localStorage.getItem('userInfo')!).groupid
        JSON.parse(localStorage.getItem('userInfo')!).store.indexOf(',') > 0 ?
          this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).store.split(',') :
          this.storeIds.push(JSON.parse(localStorage.getItem('userInfo')!).store)   
          localStorage.setItem('flag','M')    
      } else {
        if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
          this.groupId = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Preferences
          this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids.split(',')
        }
      }
      if (this.shared.common.groupsandstores.length > 0) {
        this.groupsArray = this.shared.common.groupsandstores.filter((val: any) => val.sg_id != this.shared.common.reconID);
        this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
        this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_Name : this.groupName = ''
        this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
        // // console.log(this.stores, this.groupsArray, 'Stores and Groups');
        this.getStoresandGroupsValues()
        // this.StoresData(this.ngChanges)
      }



    
    // if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
    //   this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.ustores.split(',')
    //    this.storeIds.push(JSON.parse(localStorage.getItem('userInfo')!).store) 
    // }
    // if (this.shared.common.groupsandstores.length > 0) {
    //   this.groupsArray = this.shared.common.groupsandstores.filter((val: any) => val.sg_id != this.shared.common.reconID);
    //   this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
    //   this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_Name : this.groupName = ''
    //   this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
    //   this.getStoresandGroupsValues()
    // }


    // let today = new Date();
    // let enddate = new Date(today.setDate(today.getDate() - 1));
    // this.FromDate =
    //   ('0' + (enddate.getMonth() + 1)).slice(-2) +
    //   '-01' +
    //   '-' +
    //   enddate.getFullYear();
    // this.ToDate =
    //   ('0' + (enddate.getMonth() + 1)).slice(-2) +
    //   '-' +
    //   ('0' + enddate.getDate()).slice(-2) +
    //   '-' +
    //   enddate.getFullYear();
    // this.FromDate = this.FromDate.replace(/-/g, '/');
    // this.ToDate = this.ToDate.replace(/-/g, '/');

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
    this.shared.setTitle(this.shared.common.titleName + '-Sales Manager Rankings');

    if (localStorage.getItem('Fav') != 'Y') {
      const data = {
        title: 'Sales Manager Rankings',
        path1: '',
        path2: '',
        path3: '',
        stores: this.StoreVal,
        toporbottom: this.TotalReport,
        datetype: '',
        fromdate: this.FromDate,
        todate: this.ToDate,
        groups: this.groups,
        storeorgroup: this.storeorgrp,
        dealType: this.dealType,
        saleType: this.saleType,
        financeType: this.financeType,
        count: 0
      };
      this.shared.api.SetHeaderData({
        obj: data,
      });
      this.header = [{ type: 'Bar', storeIds: this.StoreVal , storeorgroup: this.storeorgrp, dealType: this.dealType,
      saleType: this.saleType,
      financeType: this.financeType,groups:this.groups}]

      this.GetData('Rank', 'asc');
      this.setDates(this.DateType)
    } else {
      this.getFavReports();
    }
  }

  ngOnInit(): void {
    localStorage.setItem('time', this.dateType);
   

    // var curl = 'https://fbxtract.axelautomotive.com/favouritereports/GetSalesManagerRankings';
    // this.apiSrvc.logSaving(curl,{},'','Success','Sales Manager Rankings');
  }

    initializeDates(type: any) {
    let dates: any = this.setdates.setDates(type)
    this.FromDate = dates[0];
    this.ToDate = dates[1];
    localStorage.setItem('time', type);
  }



  //   setDates(type: any) {
  //   let dates: any = this.setdates.setDates(type)
  //   this.FromDate = dates[0];
  //   this.ToDate = dates[1];
  //   localStorage.setItem('time', type);
  // }
  Favreports: any = [];

  getFavReports() {
    const obj = {
      Id: localStorage.getItem('Fav_id'),
      expression: '',
    };
    this.shared.api.postmethod('favouritereports/get', obj).subscribe((res) => {
      if (res.status == 200) {
        if (res.response.length > 0) {
          this.Favreports = res.response;
          this.StoreVal = res.response[0].Fav_Report_Value;

          let dates = this.Favreports[0].Fav_Report_Value.split(',');
          this.FromDate = dates[0];
          this.ToDate = dates[1];
          localStorage.setItem('time', this.Favreports[1].Fav_Report_Value);
          this.GetData();
          localStorage.setItem('Fav', 'N');
          const data = {
            title: 'Sales Manager Rankings',
            path1: '',
            path2: '',
            path3: '',
            stores: this.StoreVal.toString(),
            toporbottom: this.TotalReport,
            datetype: '',
            fromdate: this.FromDate,
            todate: this.ToDate,
            storeorgroup: this.storeorgrp,
          };

          this.shared.api.SetHeaderData({ obj: data });
        }
      }
    });
  }
 
  
  GetData(sortdata?: any, sortstate?: any) {
    this.FIManagerData = [];
    this.shared.spinner.show();
    const obj = {
      StartDate: this.FromDate,
      EndDate: this.ToDate,
      StoreID: this.storeIds,
      Exp: sortdata,
      OrderType: sortstate,
      RankBy: this.storeorgroup,
      UserID: 0,
      SaleType: this.neworused,
      DealType: this.retailorlease,
      FinanceType: this.financetype,
    };
    let startFrom = new Date().getTime();
    // const curl = environment.apiUrl+'cavender/GetSalesManagerRankings';
    this.shared.api.postmethod(this.shared.common.routeEndpoint+'GetSalesManagerRankings', obj).subscribe(
      (res) => {
        const currentTitle = document.title;
        // this.shared.api.logSaving(curl, {}, '', res.message,currentTitle);
        if (res.status == 200) {
          if (res.response != undefined) {
            if (res.response.length > 0) {
              let resTime = (new Date().getTime() - startFrom) / 1000;
              // this.logSaving(
              //   this.solutionurl + 'cavender/GetSalesManagerRankings',
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
              let position=this.scrollCurrentposition+5
              setTimeout(() => {
                if(this.scrollcent)
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
            //   this.solutionurl + 'cavender/GetSalesManagerRankings',
            //   obj,
            //   resTime,
            //   'Error'
            // );
            // this.toast.error(res.status, '');

            this.shared.spinner.hide();
            this.NoData = true;
          }
        } else {
          // this.toast.error(res.status, '');
          this.shared.spinner.hide();
          this.NoData = true;
        }
      },
      (error) => {
        // this.toast.error('502 Bad Gate Way Error', '');
        this.shared.spinner.hide();
        this.NoData = true;
      }
    );
  }

  // GetTotalData() {
  //   this.TotalSalesPersonsData = [];
  //   const obj = {
  //     AU_ID: '69',
  //     AS_ID: this.storeIds,
  //     StartDate: this.FromDate,
  //     EndDate: this.ToDate,
  //     OrderBy: 'TR',
  //     type: 'T',
  //   };
  //   this.apiSrvc
  //     .postmethod(this.comm.routeEndpoint+'GetSalesManagerRankings', obj)
  //     .subscribe(
  //       (totalres) => {
  //         if (totalres.status == 200) {
  //           this.TotalSalesPersonsData = totalres.response.map((v) => ({
  //             ...v,
  //             Data2: [],
  //           }));
  //           this.spinner.hide();
  //           if (this.TotalSalesPersonsData.length > 0) {
  //             this.TotalSalesPersonsData.some(function (x: any) {
  //               x.data1 = 'Reports Total';
  //             });
  //             if (this.TotalReport == 'B') {
  //               this.FIManagerData.push(
  //                 this.TotalSalesPersonsData[0]
  //               );
  //             } else {
  //               this.FIManagerData.unshift(
  //                 this.TotalSalesPersonsData[0]
  //               );
  //             }
  //           }

  //           this.SalesPersonsData = [];

  //           this.SalesPersonsData = this.FIManagerData;
  //           // //console.log(this.SalesData)
  //           // this.spinner.hide()
  //           //   if(this.SalesPersonsData.length>0){
  //           //      this.SalesPersonsData.some(function(x:any){
  //           //        if(x.data1 != 'Reports Total'){
  //           //      if(x.Data2 != undefined){
  //           //       x.Data2=JSON.parse(x.Data2);
  //           //       x.Data2=x.Data2.map(v => ({ ...v, SubData:[],data2sign:'-' }))

  //           //                    }

  //           //                    if(x.Data3 != undefined){
  //           //                     x.Data3=JSON.parse(x.Data3);

  //           //                     x.Data2.forEach(val=>{

  //           //                       x.Data3.forEach(ele=>{
  //           //                         if(val.data2==ele.data2){
  //           //                           val.SubData.push(ele)
  //           //                         }
  //           //                       })
  //           //                     })
  //           //                    }
  //           //        }
  //           //     x.Dealer ='+';
  //           //     return false;
  //           //   });
  //           // }
  //           if (this.SalesPersonsData.length == 0) {
  //             this.NoData = true;
  //           } else {
  //             this.NoData = false;
  //           }
  //         } else {
  //     
  //         }
  //       },
  //       (error) => {
  //         //console.log(error);
  //       }
  //     );
  // }

  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    }
    return false;
  }
  // reportOpen(temp: any) {


  //   this.ngbmodalActive = this.ngbmodal.open(temp, {
  //     size: 'xl',
  //     backdrop: 'static',
  //   });
  // }
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
  //     if( this.storeorgrp == 'G' && (col_Name != 'Rank' && col_Name != 'fimanager' && col_Name != 'StoreName'  )){
  //       this.columnState = 'desc';
  //       this.columnName = col_Name;
  //       this.GetData(this.columnName, this.columnState);
  //     }else{
  //       this.columnState = 'asc';
  //       this.columnName = col_Name;
  //       this.GetData(this.columnName, this.columnState);
  //     }
    
  //   }
  //   // //console.log(this.columnName, this.columnState);
  // }
  tabClick(col_Name: any) {

    // First click on a column
    if (this.columnName !== col_Name) {
      this.columnName = col_Name;
      this.columnState = 'asc';   // 👈 start with ASC on first click
    }
    // Clicking same column again → toggle
    else {
      this.columnState = this.columnState === 'asc' ? 'desc' : 'asc';
    }
  
    this.GetData(this.columnName, this.columnState);
  }
  ngAfterViewInit(): void {
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.shared.common.pageName == 'Sales Manager Rankings') {
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
      // //console.log(res);
      if (res.obj.Module == 'Sales Manager Rankings') {
        document.getElementById('report')?.click()
      }
    });
    this.shared.api.GetReports().subscribe((data) => {
      if (data.obj.Reference == 'Sales Manager Rankings') {
        if(data.obj.header == undefined){
        this.StoreVal = data.obj.storeValues;
        this.TotalReport = data.obj.TotalReport;
        this.storeorgrp = data.obj.storeorgroup;
        this.groups=data.obj.groups
        if (data.obj.FromDate != undefined && data.obj.ToDate != undefined) {
          this.FromDate = data.obj.FromDate;
          this.ToDate = data.obj.ToDate;
          this.storeIds = data.obj.storeValues;
          this.dateType = data.obj.dateType;
          this.dealType = data.obj.dealType;
          this.saleType = data.obj.saleType;
          this.financeType = data.obj.financeType;
          this.storeorgrp == 'G' && (this.columnName != 'Rank' && this.columnName != 'fimanager' &&this.columnName != 'StoreName'  ) ? this.columnState='asc' :'';
          this.GetData(this.columnName, this.columnState);
        } else {
          this.FromDate = data.obj.FromDate;
          this.ToDate = data.obj.ToDate;
          this.storeIds = data.obj.storeValues;
          this.dateType = data.obj.dateType;
          this.dealType = data.obj.dealType;
          this.saleType = data.obj.saleType;
          this.financeType = data.obj.financeType;
          this.storeorgrp == 'G' && (this.columnName != 'Rank' && this.columnName != 'fimanager' &&this.columnName != 'StoreName'  ) ? this.columnState='asc' :'';

          this.GetData(this.columnName, this.columnState);
        }
      }
      else{
        if(data.obj.header == 'Yes'){
          this.StoreVal = data.obj.storeValues;
        //console.log(this.StoreVal);
        this.GetData(this.columnName, this.columnState);
        }            
      }
        const headerdata = {
          title: 'Sales Manager Rankings',
          path1: '',
          path2: '',
          path3: '',
          stores: this.StoreVal,
          toporbottom:  this.TotalReport,
          datetype: this.dateType,
          fromdate:    this.FromDate,
          todate: this.ToDate,
          groups: this.groups,
          storeorgroup: this.storeorgrp,
          dealType: this.dealType,
          saleType: this.saleType,
          financeType: this.financeType,
        };
        this.shared.api.SetHeaderData({
          obj: headerdata,
        });
        this.header = [{ type: 'Bar', storeIds: this.StoreVal,storeorgroup: this.storeorgrp, dealType: this.dealType,
        saleType: this.saleType,
        financeType: this.financeType ,groups:this.groups}]
      }
    });
    this.shared.api.getExportToExcelAllReports().subscribe((res) => {
      this.FIMstate = res.obj.state;
      if (res.obj.title == 'Sales Manager Rankings') {
        if (res.obj.state == true) {
          this.exportToExcel();
        }
      }
    });

    this.shared.api.getExportToEmailPDFAllReports().subscribe((res) => {
      if (res.obj.title == 'Sales Manager Rankings') {
        if (res.obj.stateEmailPdf == true) {
          // this.sendEmailData(res.obj.Email, res.obj.notes, res.obj.from);
        }
      }
    });
    this.shared.api.getExportToPrintAllReports().subscribe((res) => {
      if (res.obj.title == 'Sales Manager Rankings') {
        if (res.obj.statePrint == true) {
          this.GetPrintData();
        }
      }
    });
    this.shared.api.getExportToPDFAllReports().subscribe((res) => {
      if (res.obj.title == 'Sales Manager Rankings') {
        if (res.obj.statePDF == true) {
          // this.generatePDF();
        }
      }
    });
  }

  // openDetails(Item) {
  //   this.CompleteComponentState = false;
  //   const DetailsSalesPeron = this.ngbmodal.open(SalespersonsDealsComponent, {
  //     // size:'xl',
  //     backdrop: 'static',
  //   });
  //   DetailsSalesPeron.componentInstance.Dealdetails = Item;
  //   DetailsSalesPeron.result.then(
  //     (data) => {},
  //     (reason) => {
  //       // on dismiss
  //       this.CompleteComponentState = true;
  //     }
  //   );
  // }


  Scrollpercent: any = 0;
  scrollCurrentposition:any=0
  @ViewChild('scrollcent') scrollcent!: ElementRef;

  updateVerticalScroll(event: any): void {

    this.scrollCurrentposition=event.target.scrollTop
    const scrollDemo = document.querySelector('#scrollcent') as HTMLElement;
    this.Scrollpercent = Math.round(
      (event.target.scrollTop /
        (event.target.scrollHeight - scrollDemo.clientHeight)) *
      100
    );
  }

  selBlock: any;
  commentopen(item: any, i: any, slblock: any = '') {
    this.index = '';
    //console.log('Selected Obj :', item);
    //return
    this.selBlock = slblock + i.toString();
    this.index = i.toString();
    this.commentobj = {
      TYPE: item.fimanager,
      NAME: item.fimanager,
      STORES: item.StoreName,
      STORENAME: item.StoreName,
      Month: '',
      ModuleId: '24',
      ModuleRef: 'FIMR',
      state: 1,
      indexval: i,
    };
  }
    //----------REPORTS---------------//
    activePopover: number = -1;
    storeorgroup: any = ['G'];
    retailorlease: any = ['Retail', 'Lease', 'Misc', 'Special Order'];
    financetype: any = ['Finance', 'Cash', 'Lease'];
    neworused: string[] = ['New', 'Used'];
    selectedGroups: any[] = [];
    AllStores: boolean = true;
    AllGroups: boolean = true;
      toporbottom: any = ['T'];
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
    allgroups() {
      this.AllGroups = !this.AllGroups;
      this.selectedGroups = this.AllGroups ? this.groups.map((g: { sg_id: any; }) => g.sg_id) : [];
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
    updatedDates(data: any) {
      console.log(data);
      this.FromDate = data.FromDate;
      this.ToDate = data.ToDate;
      this.DateType = data.DateType;
      this.displaytime = data.DisplayTime
    }
    setDates(type: any) {
      // localStorage.setItem('time', type);
      // this.datevaluetype=
      console.log(type);
  
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
   
      // Deal type & status
      multipleorsingle(block: any, e: any) {
        if (block == 'NU') {
          const index = this.neworused.findIndex((i: any) => i == e);
          if (index >= 0) {
            this.neworused.splice(index, 1);
          } else {
            this.neworused.push(e);
          }
        }
        
        if (block == 'RL') {
          const index = this.retailorlease.findIndex((i: any) => i == e);
          if (index >= 0) {
            this.retailorlease.splice(index, 1);
          } else {
            this.retailorlease.push(e);
          }
        }
        if (block == 'DS') {
          const index = this.financetype.findIndex((i: any) => i == e);
          if (index >= 0) {
            this.financetype.splice(index, 1);
          } else {
            this.financetype.push(e);
          }
        }
      
      }
  // groupstate: boolean = false
  // individualgroups(e: any) {
  //   this.groupstate = true
  //   const index = this.selectedGroups.findIndex((i: any) => i == e.sg_id);
  //   if (index >= 0) {
  //     // this.selectedGroups.splice(index, 1);
  //     // this.AllGroups = false;
  //     // this.getGroupBaseStores(this.selectedGroups.toString())
  //   } else {
  //     this.selectedGroups = []
  //     this.selectedGroups.push(e.sg_id);
  //     this.groupName = this.groups.filter((val: any) => val.sg_id == this.selectedGroups[0])[0].sg_name;
      
  //     if(e.sg_id==1){
  //       this.storeorgroup="S"
  //     }
  //     else{
  //       this.storeorgroup="G"
  //     }
  //     this.getGroupBaseStores(this.selectedGroups.toString())
  //     if (this.selectedGroups.length == this.groups.length) {
  //       this.AllGroups = true;
  //     } else {
  //       this.AllGroups = false;
  //     }
  //   }
  // }
  
  storeorgroups(block: any, e: any) {
    // if (block == 'TB') {
      this.storeorgroup = [];
      this.storeorgroup.push(e)
    // }
  }
  viewreport() {
    this.activePopover = -1;
    console.log('Apply clicked with:', {
      FromDate: this.FromDate,
      ToDate: this.ToDate,
      Stores: this.storeIds,
      Groups: this.selectedGroups,
      StoreOrGroup: this.storeorgroup,
      SaleType: this.retailorlease,
      financetype: this.financetype,

      dealType: this.dealType,
      saleType: this.saleType,
      financeType: this.financeType
    });


    
    if (!this.storeIds || this.storeIds.length === 0) {

      this.toast.show('Please select at least one store', 'warning', 'Warning');
      return; 
    } else 
    if (this.retailorlease.length == 0) {
     
      this.toast.show('Please select any one Deal Type', 'warning', 'Warning');
    }
    else if (this.neworused.length == 0) {
    
      this.toast.show('Please select atleast one Sale Type', 'warning', 'Warning');
    }
    else if (this.financetype.length == 0) {
   
      this.toast.show('Please select atleast one Finance Type', 'warning', 'Warning');
    }
    else {
      const data = {
        Reference: 'F&I Manager Ranking',
        FromDate: this.FromDate,
        ToDate: this.ToDate,
        TotalReport: this.toporbottom[0],
        storeValues: this.storeIds.toString(),
        // == '' ? '0': this.selectedstorevalues.toString(),
        dateType: this.DateType,
        groups: this.selectedGroups.toString(),
        storeorgroup: this.storeorgroup.toString(),
        saleType: this.retailorlease.toString(),
        financetype: this.financetype,
        dealType: this.dealType,
        // saleType: this.saleType
      };
      this.shared.api.SetReports({
        obj: data,
      });
      this.close();
      this.GetData(this.columnName, this.columnState);


    }
  }
  close() {
    this.shared.ngbmodal.dismissAll();
  }

    //   setDates(type: any) {
    //   // localStorage.setItem('time', type);
    //   // this.datevaluetype=
    //   console.log(type);
  
    //   this.displaytime = 'Time Frame (' + this.Dates.Types.filter((val: any) => val.code == type)[0].name + ')';
    //   this.maxDate = new Date();
    //   this.minDate = new Date();
    //   this.minDate.setFullYear(this.maxDate.getFullYear() - 3);
    //   this.maxDate.setDate(this.maxDate.getDate());
    //   this.Dates.FromDate = this.FromDate;
    //   this.Dates.ToDate = this.ToDate;
    //   this.Dates.MinDate = this.minDate;
    //   this.Dates.MaxDate = this.maxDate;
    //   this.Dates.DateType = this.DateType;
    //   this.Dates.DisplayTime = this.displaytime;
    // }
    // const DetailsSF = this.ngbmodal.open(CommentsComponent, {
    //   size: 'xl',
    //   backdrop: 'static',
    // });
    // DetailsSF.componentInstance.SFComments = {
    //   TYPE: item.LABLEVAL,
    //   NAME: item.LABLE,
    //   STORES: this.selectedstorevalues,
    //   LatestDate: this.Month,
    //   STORENAME: this.selectedstorename,
    //   Month: this.Month,
    //   ModuleId: '66',
    //   ModuleRef: 'SF',
    // };
    // DetailsSF.result.then(
    //   (data) => {},
    //   (reason) => {
  
    //     // // on dismiss
    //     // const Data = {
    //     //   state: true,
    //     // };
    //     // this.apiSrvc.setBackgroundstate({ obj: Data });
    //     this.GetData();
    //   }
    // );




  index = '';
  commentobj: any = {};

  // addcmt(data: any) {
  //   if (data == 'A') {
  //     this.index = '';
  //     const DetailsSF = this.ngbmodal.open({
  //       size: 'xl',
  //       backdrop: 'static',
  //     });
  //     // myObject['skillItem2'] = 15;
  //     this.commentobj['state'] = 0;
  //     (DetailsSF.componentInstance.SFComments = this.commentobj),
  //       DetailsSF.result.then(
  //         (data) => {
  //           // //console.log(data);
  //         },
  //         (reason) => {
  //           // //console.log(reason);

  //           if (reason == 'O') {
  //             this.commentobj['state'] = 1;
  //             this.index = this.commentobj['indexval'];
  //           } else {
  //             this.commentobj['state'] = 1;
  //             this.index = this.commentobj['indexval'];

  //             this.GetData(this.columnName, this.columnState);

  //           }
  //           // // on dismiss

  //           // const Data = {
  //           //   state: true,
  //           // };
  //           // this.apiSrvc.setBackgroundstate({ obj: Data });
  //           // this.GetData();
  //         }
  //       );
  //   }
  //   if (data == 'AD') {

  //     this.GetData(this.columnName, this.columnState);

  //   }
  // }
  // close(data: any) {
  //   // //console.log(data);
  //   this.index = '';
  // }
  // ExcelStoreNames: any = [];
  // exportToExcel() {
  //   let storeNames: any = [];
  //   let store = this.StoreVal.split(',');
  //   // const obj = {
  //   //   id: this.groups,
  //   //   userid: JSON.parse(localStorage.getItem('UserDetails')!).userid,
  //   // };
  //   // this.apiSrvc
  //   //   .postmethodOne('cavender/GetStoresbyGroupuserid', obj)
  //   //   .subscribe((res: any) => {
  //       storeNames =this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groups)[0].Stores
  //       .filter((item: any) =>
  //         store.some((cat: any) => cat === item.ID.toString())
  //       );
  //       // //console.log(store,res.response.length);
        
  //       if (store.length ==this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groups)[0].Stores
  //       .length) {
  //         this.ExcelStoreNames = 'All Stores'
  //       } else {
  //         this.ExcelStoreNames = storeNames.map(function (a: any) {
  //           return a.storename;
  //         });
  //       }
  //   const FIManagerData = this.FIManagerData.map((_arrayElement: any) =>
  //     Object.assign({}, _arrayElement)
  //   );
  //   // //console.log(FIManagerData);

  //   const workbook = new Workbook();
  //   const worksheet = workbook.addWorksheet('Sales Manager Rankings');
  //   worksheet.views = [
  //     {
  //       state: 'frozen',
  //       ySplit: 14, // Number of rows to freeze (2 means the first two rows are frozen)
  //       topLeftCell: 'A15', // Specify the cell to start freezing from (in this case, the third row)
  //       showGridLines: false,
  //     },
  //   ];
  //   worksheet.addRow('');
  //   const titleRow = worksheet.addRow(['Sales Manager Rankings']);
  //   titleRow.eachCell((cell, number) => {
  //     cell.alignment = { indent: 1, vertical: 'middle', horizontal: 'left' };
  //   });
  //   titleRow.font = { name: 'Arial', family: 4, size: 12, bold: true };
  //   titleRow.worksheet.mergeCells('A2', 'D2');
  //   worksheet.addRow('');
  //   const DateToday = this.shared.datePipe.transform(
  //     new Date(),
  //     'MM.dd.yyyy h:mm:ss a'
  //   );
  //   const DATE_EXTENSION = this.shared.datePipe.transform(
  //     new Date(),
  //     'MMddyyyy'
  //   );
  //   worksheet.addRow([DateToday]).font = { name: 'Arial', family: 4, size: 9 };
  //   worksheet.getCell('A4').alignment = { vertical: 'middle', horizontal: 'left',indent:1};

  //   const ReportFilter = worksheet.addRow(['Report Controls :']);
  //   ReportFilter.font = { name: 'Arial', family: 4, size: 10, bold: true };
  //   ReportFilter.alignment = { vertical: 'middle', horizontal: 'left',indent:1};

  //   const Timeframe = worksheet.addRow(['Timeframe :']);
  //   Timeframe.alignment = { vertical: 'middle', horizontal: 'left',indent:1};
  //   Timeframe.getCell(1).font = {
  //     name: 'Arial',
  //     family: 4,
  //     size: 9,
  //     bold: true,
  //   };
  //   const timeframe = worksheet.getCell('B6');
  //   timeframe.value = this.FromDate + ' to ' + this.ToDate;
  //   timeframe.font = { name: 'Arial', family: 4, size: 9 };
  //   timeframe.alignment = { vertical: 'middle', horizontal: 'left'};

  //   const Groups = worksheet.getCell('A7');
  //   Groups.value = 'Group :';
  //   Groups.alignment = { vertical: 'middle', horizontal: 'left',indent:1};
  //   Groups.font = { name: 'Arial', family: 4, size: 9,bold:true};
  //   const groups = worksheet.getCell('B7');
  //   groups.value =
  //   this.shared.common.groupsandstores.filter((val: any) => val.sg_id == this.groups.toString())[0].sg_name;
  //   groups.font = { name: 'Arial', family: 4, size: 9};
  //   groups.alignment = { vertical: 'middle', horizontal: 'left'};
  //   worksheet.mergeCells('B8', 'K10');
  //   const Stores = worksheet.getCell('A8');
  //   Stores.value = 'Stores :'
  //   Stores.alignment = { vertical: 'middle', horizontal: 'left',indent:1};
  //   const stores = worksheet.getCell('B8');
  //   stores.value = this.ExcelStoreNames == 0
  //   ? '-'
  //   : this.ExcelStoreNames == null
  //   ? '-'
  //   : this.ExcelStoreNames.toString().replaceAll(',', ', ');
  //   stores.font = { name: 'Arial', family: 4, size: 9 };
  //   stores.alignment = { vertical: 'top', horizontal: 'left',wrapText:true};
  //   Stores.font = {
  //     name: 'Arial',
  //     family: 4,
  //     size: 9,
  //     bold: true,
  //   };

  //   const Filter = worksheet.getCell('A11');
  //   Filter.value ='Filter :';
  //   Filter.alignment = { vertical: 'middle', horizontal: 'left',indent:1};
  //   Filter.font = { name: 'Arial', family: 4, size: 9,bold:true};
  //   const filter = worksheet.getCell('B11');
  //   filter.value = this.storeorgrp == 'S' ? 'Rank By Store' : 'Rank By group';
  //   filter.font = { name: 'Arial', family: 4, size: 9 };
  //   filter.alignment = { vertical: 'middle', horizontal: 'left' };
  //   worksheet.addRow('');
  //   worksheet.mergeCells('A13', 'C13');
  //   let MTD = worksheet.getCell('A13');
  //   MTD.value = 'MTD';
  //   MTD.alignment = { vertical: 'middle', horizontal: 'center' };
  //   MTD.font = {
  //     name: 'Arial',
  //     family: 4,
  //     size: 9,
  //     bold: true,
  //     color: { argb: 'FFFFFF' },
  //   };
  //   MTD.fill = {
  //     type: 'pattern',
  //     pattern: 'solid',
  //     fgColor: { argb: '2a91f0' },
  //     bgColor: { argb: 'FF0000FF' },
  //   };
  //   MTD.border = { right: { style: 'thin' } };

  //   worksheet.mergeCells('D13', 'H13');
  //   let UnitCredit = worksheet.getCell('D13');
  //   UnitCredit.value = 'Unit Credit';
  //   UnitCredit.alignment = { vertical: 'middle', horizontal: 'center' };
  //   UnitCredit.font = {
  //     name: 'Arial',
  //     family: 4,
  //     size: 9,
  //     bold: true,
  //     color: { argb: 'FFFFFF' },
  //   };
  //   UnitCredit.fill = {
  //     type: 'pattern',
  //     pattern: 'solid',
  //     fgColor: { argb: '2a91f0' },
  //     bgColor: { argb: 'FF0000FF' },
  //   };
  //   UnitCredit.border = { right: { style: 'thin' } };

  //   worksheet.mergeCells('I13', 'L13');
  //   let FrontGross = worksheet.getCell('I13');
  //   FrontGross.value = 'Gross';
  //   FrontGross.alignment = { vertical: 'middle', horizontal: 'center' };
  //   FrontGross.font = {
  //     name: 'Arial',
  //     family: 4,
  //     size: 9,
  //     bold: true,
  //     color: { argb: 'FFFFFF' },
  //   };
  //   FrontGross.fill = {
  //     type: 'pattern',
  //     pattern: 'solid',
  //     fgColor: { argb: '2a91f0' },
  //     bgColor: { argb: 'FF0000FF' },
  //   };
  //   FrontGross.border = { right: { style: 'thin' } };

  //   let Headings = [
  //     'Rank',
  //     'Sales Manager',
  //     'Store Name',
  //     'New',
  //     'Used',
  //     'Total',
  //     'Pace',
  //     '90 Day Avg',
  //     'Gross',
  //     'Pace',
  //     'PVR',
  //       '90 Day Avg',
  //   ];
  //   const headerRow = worksheet.addRow(Headings);
  //   headerRow.font = {
  //     name: 'Arial',
  //     family: 4,
  //     size: 9,
  //     bold: false,
  //     color: { argb: 'FFFFFF' },
  //   };
  //   headerRow.alignment = { indent: 1, vertical: 'middle', horizontal: 'center' };
  //   headerRow.eachCell((cell, number) => {
  //     cell.fill = {
  //       type: 'pattern',
  //       pattern: 'solid',
  //       fgColor: { argb: '788494' },
  //       bgColor: { argb: 'FF0000FF' },
  //     };
  //     cell.border = { right: { style: 'thin' } };
  //     cell.alignment = { vertical: 'middle', horizontal: 'center' };
  //   });
  //   for (const d of FIManagerData) {
  //     const Data = worksheet.addRow([
  //       d.Rank == '' ? '-' : d.Rank == null ? '-' : d.Rank,
  //       d.SalesManager == '' ? '-' : d.SalesManager == null ? '-' : d.SalesManager,
  //       d.StoreName == '' ? '-' : d.StoreName == null ? '-' : d.StoreName,
  //       d.New == '' ? '-' : d.New == null ? '-' : d.New,
  //       d.Used == '' ? '-' : d.Used == null ? '-' : d.Used,
  //       d.Total == '' ? '-' : d.Total == null ? '-' : d.Total,
  //       d.Pace == '' ? '-' : d.Pace == null ? '-' : parseInt(d.Pace),
  //       d.UnitsDayAvg == '' ? '-' : d.UnitsDayAvg == null ? '-' : parseInt(d.UnitsDayAvg),

    
  //       d.BackGross == '' ? '-' : d.BackGross == null ? '-' : d.BackGross,
  //       d.BackGrossPace == ''
  //         ? '-'
  //         : d.BackGrossPace == null
  //         ? '-'
  //         : d.BackGrossPace,
  //       d.BackGrossPvr == ''
  //         ? '-'
  //         : d.BackGrossPvr == null
  //         ? '-'
  //         : d.BackGrossPvr,
  //       d.GrossDayAvg == '' ? '-' : d.GrossDayAvg == null ? '-' : parseInt(d.GrossDayAvg),

  //     ]);
  //     // Data1.outlineLevel = 1; // Grouping level 1
  //     Data.font = { name: 'Arial', family: 4, size: 8 };
  //     Data.alignment = { vertical: 'middle', horizontal: 'center' };
  //     Data.getCell(1).alignment = {
  //       indent: 1,
  //       vertical: 'middle',
  //       horizontal: 'center'
  //     };
  //     if (Data.number % 2) {
  //       Data.eachCell((cell, number) => {
  //         cell.fill = {
  //           type: 'pattern',
  //           pattern: 'solid',
  //           fgColor: { argb: 'e5e5e5' },
  //           bgColor: { argb: 'FF0000FF' },
  //         };
  //       });
  //     }
  //     Data.eachCell((cell, number) => {
  //       cell.border = { right: { style: 'thin' } };
  //       if (number > 3 && number < 8) {
  //         cell.numFmt = '#,##0';
  //         cell.alignment = { vertical: 'middle', horizontal: 'center', indent: 1 };
  //       }
  //       if (number > 8 && number < 13) {
  //         cell.numFmt = '$#,##0';
  //         cell.alignment = { vertical: 'middle', horizontal: 'center', indent: 1 };
  //       }
  //       if (number == 2 || number == 3) {
  //         cell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
  //       }
  //     });
  //   }
  //   worksheet.getColumn(1).width = 15;
  //   worksheet.getColumn(2).width = 30;
  //   worksheet.getColumn(3).width = 30;
  //   worksheet.getColumn(4).width = 15;
  //   worksheet.getColumn(5).width = 15;
  //   worksheet.getColumn(6).width = 15;
  //   worksheet.getColumn(7).width = 15;
  //   worksheet.getColumn(8).width = 15;
  //   worksheet.getColumn(9).width = 15;
  //   worksheet.getColumn(10).width = 15;
  //   worksheet.getColumn(11).width = 15;
  //   worksheet.getColumn(12).width = 15;


  //   worksheet.addRow([]);
  //   workbook.xlsx.writeBuffer().then((data: any) => {
  //     const blob = new Blob([data], {
  //       type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  //     });
  //     FileSaver.saveAs(blob, 'Sales Manager Ranking_'+DATE_EXTENSION+ EXCEL_EXTENSION);
  //   });
  // // });
  // }
  ExcelStoreNames: any = []
  exportToExcel(): void {
    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('Sales Manager Rankings');
    const title = worksheet.addRow(['Sales Manager Rankings']);
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
      let store: any = [];
      this.storeIds && this.storeIds.toString().indexOf(',') > 0 ? store = this.storeIds.toString().split(',') : store.push(this.storeIds)
      console.log(this.storeIds, this.groups, '...............................');
      // storeNames = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groups)[0].Stores.filter((item: any) => store.includes(item.ID.toString()));
      // if (store.length == this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groups)[0].Stores.length) { 
      //   this.ExcelStoreNames = 'All Stores' }
      // else { 
        this.ExcelStoreNames = storeNames.map(function (a: any) { return a.storename; });
      //  }
  
      let storeValue = '';

      const selectedStoreIds: string[] =
        this.storeIds && this.storeIds.length
          ? this.storeIds.map((id: any) => id.toString())
          : [];
      
      const allStores: any[] = Array.isArray(this.stores) ? this.stores : [];
      
      // ✅ Bind ONLY selected store names
      storeValue = allStores
        .filter((s: any) => selectedStoreIds.includes(s.ID.toString()))
        .map((s: any) => s.storename.trim())
        .filter(Boolean)
        .join(', ');
      
      // ✅ Final fallback (safety)
      if (!storeValue && selectedStoreIds.length) {
        storeValue = selectedStoreIds.join(', ');
      }
      
          const filters = [
            { name: 'Store:', values: storeValue },
      { name: 'Time Frame:', values: `${formattedFromDate} to ${formattedToDate}` },
      { name: 'Rank By:', values: this.storeorgroup == 'S' ? 'Store' : 'Group' },
      // { name: 'New/Used:', values: this.neworused || 'All' },
      { name: 'Deal Type:', values: this.retailorlease || 'All' },
      // { name: 'Deal Status:', values: this.dealStatus || 'All' },
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
      'Rank',
      'Sales Manager',
      'Store Name',
      'New',
      'Used',
      'Total',
      'Pace',
      '90 Day Avg',
      'Gross',
      'Pace',
      'PVR',
        '90 Day Avg',
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
      'Rank', 'SalesManager', 'StoreName',
      'New', 'Used',
      'Total', 'Pace', 'UnitDayAvg', 'BackGross', 'BackGrossPace', 'BackGrossPvr', 'GrossDayAvg'
    ];
  
    const currencyFields = ['GroBackGrossss', 'BackGrossPace', 'BackGrossPvr', 'GrossDayAvg'];
  
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
      this.shared.exportToExcel(workbook, 'Sales Manager Rankings');
    });
  }
  scrollPosition = 0;

  getScrollPosition(event: any): void {
    this.scrollPosition = event.target.scrollLeft ;
    console.log(this.scrollPosition,event.target.scrollTop);
    
  }
//   GetPrintData() {
//     let printContents, popupWin;
//     printContents = document.getElementById('F&IManagerRankings')!.innerHTML;
//     popupWin = window.open('', '_blank', 'top=0,left=0,height=100%,width=auto');
//     popupWin!.document.open();
//     popupWin!.document.write(`
//       <html>
//       <head>
//       <title>Sales Manager Rankings</title>
//       <style>
     
// @media print {
//   body {-webkit-print-color-adjust: exact;}
//   @page {
//     size: landscape;
//     size: A4 landscape;
//   }
// }
//       @font-face {
//        font-family: 'GothamBookRegular';
//        src: url('assets/fonts/Gotham\ Book\ Regular.otf') format('otf'), /* Chrome 6+, Firefox 3.6+, IE 9+, Safari 5.1+ */
//             url('assets/fonts/Gotham\ Book\ Regular.otf') format('opentype'); /* Chrome 4+, Firefox 3.5, Opera 10+, Safari 3—5 */
//      }
//      @font-face {
//        font-family: 'Roboto';
//        src: url('assets/fonts/Roboto-Regular.ttf') format('ttf'), /* Chrome 6+, Firefox 3.6+, IE 9+, Safari 5.1+ */
//             url('assets/fonts/Roboto-Regular.ttf') format('truetype'); /* Chrome 4+, Firefox 3.5, Opera 10+, Safari 3—5 */
//      }
//      @font-face {
//        font-family: 'RobotoBold';
//        src: url('assets/fonts/Roboto-Bold.ttf') format('ttf'), /* Chrome 6+, Firefox 3.6+, IE 9+, Safari 5.1+ */
//             url('assets/fonts/Roboto-Bold.ttf') format('truetype'); /* Chrome 4+, Firefox 3.5, Opera 10+, Safari 3—5 */
//      }
//      .justify-content-between {
//       justify-content: space-between !important;
//   }
//   .d-flex {
//       display: flex !important;
//   }
//     .performance-scorecard .table>:not(:first-child) {
//       border-top: 0px solid #ffa51a
//     }
//     .performance-scorecard .table {
//       text-align: center;
//       text-transform: capitalize;
//       border: transparent;

//       width: 100%;
//     }
//     .performance-scorecard .table th,
//     .performance-scorecard .table td{
//       white-space: nowrap;
//       vertical-align: top;
//     }
//     .performance-scorecard .table th:first-child {
//       position: sticky;
//       left: 0;
//       z-index: 1;
//       // background-color: #337ab7;
//    }
//    .performance-scorecard .table td:first-child {
//     position: sticky;
//     left: 0;
//     z-index: 1;
//     // background-color: #337ab7;
//   }
//   .performance-scorecard .table tr:nth-child(odd) {
//     // background-color: #e9ecef;
//     background-color: #ffffff;

//   }
//   .performance-scorecard .table  tr:nth-child(even) {
//     background-color: #ffffff;
//   }
//   .performance-scorecard .table .spacer {
//     // width: 50px !important;
//     background-color: #cfd6de !important;
//     border-left: 1px solid #cfd6de !important;
//     border-bottom: 1px solid #cfd6de !important;
//     border-top: 1px solid #cfd6de !important;
//   }
//   .performance-scorecard .table .hidden {
//     display: none !important;
//   }
//   .performance-scorecard .table .bdr-rt {
//     border-right: 1px solid #abd0ec;
//   }
//   .performance-scorecard .table thead {
//     position: sticky;
//     top: 0;
//     z-index: 99;
//     font-family: 'FaktPro-Bold';
//     font-size: 0.8rem;
//   }
//   .performance-scorecard .table thead th {
//     padding: 5px 10px;
//     margin: 0px;
//   }
//   .performance-scorecard .table thead .bdr-btm {
//     border-bottom: #005fa3;
//   }
//   .performance-scorecard .table thead tr:nth-child(1) {
//     background-color: #fff !important;
//     color: #000;
//     // color: #fff;
//     text-transform: uppercase;
//     border-bottom: #cfd6de;
//   }
//   .performance-scorecard .table thead tr:nth-child(2) {
//     background-color: #337ab7 !important;
//     color: #fff;
//     text-transform: uppercase;
//     border-bottom: #cfd6de;
//     box-shadow: inset 0 1px 0 0 #cfd6de;
//   }
//   .performance-scorecard .table thead  tr:nth-child(3) {

//     background-color: #fff !important;
//     color: #000;
//     text-transform: uppercase;
//     border-bottom: #cfd6de;
//     box-shadow: inset 0 1px 0 0 #cfd6de;
//   }
//   .performance-scorecard .table thead  tr:nth-child(3) th :nth-child(1) {
//     background-color: #337ab7 !important;
//     color: #fff;
//   }
//   .performance-scorecard .table tbody {
//     font-family: 'FaktPro-Normal';
//     font-size: .9rem;
//   }
//   .performance-scorecard .table tbody td {
//     padding: 2px 10px;
//     margin: 0px;
//     border: 1px solid #cfd6de
//   }
//   .performance-scorecard .table tbody tr {
//     border-bottom: 1px solid #37a6f8;
//     border-left: 1px solid #37a6f8
//   }
//   .performance-scorecard .table tbody td:first-child {
//     text-align: start;
//     box-shadow: inset -1px 0 0 0 #cfd6de;
//   }
//   .performance-scorecard .table tbody tr td:not(:first-child) {
//     text-align: right !important;

//   }
//   .performance-scorecard .table tbody .sub-title {
//     font-size: .8rem !important;
//   }
//   .performance-scorecard .table tbody .sub-subtitle{
//     font-size: .7rem !important;
//   }
//   .performance-scorecard .table tbody td:nth-child(2){ 
//     padding: 2px 10px;
//     margin: 0px;
//   }
//   .performance-scorecard .table tbody .text-bold {
//     font-family: 'FaktPro-Bold';
//   }
//   .performance-scorecard .table tbody .darkred-bg {
//     background-color: #282828 !important;
//     color: #fff;
//   }
//   .performance-scorecard .table tbody .lightblue-bg {
//     background-color: #646e7a !important;
//     color: #fff;
//   }
//   .performance-scorecard .table tbody .gold-bg {
//     background-color: #ffa51a;
//     color: #fff;
//   }


//       </style>
//     </head>
//   <body onload="window.print();window.close()">${printContents}</body>
//     </html>`);
//     popupWin!.document.close();
//   }
  GetPrintData() {
    window.print();
   }

}
