import { Component, OnInit, ViewChild, ElementRef, HostListener } from '@angular/core';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { Setdates } from '../../../../Core/Providers/SetDates/setdates';
import { BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { common } from '../../../../common';
import { DateRangePicker } from '../../../../CommonFilters/date-range-picker/date-range-picker';
import { Subscription } from 'rxjs';
import { Stores } from '../../../../CommonFilters/stores/stores';
import { NgbModule } from '@ng-bootstrap/ng-bootstrap';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [SharedModule, BsDatepickerModule, DateRangePicker,Stores,NgbModule],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
})
export class Dashboard {
  SalesPersonsData: any = [];
  ServiceAdvisorData: any = [];

  FromDate: any = '';
  ToDate: any = '';
  NoData: boolean = false;
 
 
  storeIds: any = '0';
  CompleteComponentState: boolean = true;
  dateType: any = 'MTD';
  Paytype: any = ['C', 'W', 'I', 'S', 'M']; 

  // solutionurl: any = environment.apiUrl;
  LogCount = 1;
  groups: any = 0;
  StoreVal: any = '0';
  LaborTypeVal: any = ''
  columnName: any = 'Rank';
  columnState: any = 'asc';
  storeorgrp: any = 'G';
  zeroro: any = 'E';
  LaborState: any = 'S';
  otherstoreid: any = '';
  selectedotherstoreids: any = '';

  reportOpenSub!: Subscription;
  reportGetting!: Subscription;
  Pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;
  // FromDate: any = '';
  // ToDate: any = '';
  minDate!: Date;
  maxDate!: Date;
  DateType: any = 'MTD';
  displaytime: any = '';
  stores: any = []
  groupsArray: any = [];
  storename: any = ''
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  groupId: any = 0;

  storesFilterData: any = {
    'groupsArray': this.groupsArray, 'groupId': this.groupId, 'storesArray': this.stores, 'storeids': '1', 'type': 'M', 'others': 'N',
    'groupName': this.groupName, 'storename': this.storename, storecount: null, 'storedisplayname': this.storedisplayname
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

  constructor(public shared: Sharedservice, public setdates: Setdates, private comm: common,private toast: ToastService,  ) {
    this.initializeDates(this.DateType)
    let today = new Date();
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

    let enddate = new Date(today.setDate(today.getDate() - 1));
    this.FromDate =
      ('0' + (enddate.getMonth() + 1)).slice(-2) +
      '-01' +
      '-' +
      enddate.getFullYear();
    // let dt = new Date(today.setDate(today.getDate() - 1));   
    // dt.setMonth(dt.getMonth() - 3);

    // this.FromDate = this.datepipe.transform(dt,'MM-dd-yyyy')
    this.ToDate =
      ('0' + (enddate.getMonth() + 1)).slice(-2) +
      '-' +
      ('0' + enddate.getDate()).slice(-2) +
      '-' +
      enddate.getFullYear();
    this.FromDate = this.FromDate.replace(/-/g, '/');
    this.ToDate = this.ToDate.replace(/-/g, '/');
    this.shared.setTitle(this.shared.common.titleName + '-Service Advisor Rankings');
    const data = {
      title: 'Service Advisor Rankings',  
      stores:   this.storeIds.toString(),
      labortype: this.LaborTypeVal.toString(),
      laborstate: this.LaborState,
    
      datetype: 'MTD',
      fromdate: this.FromDate,
      todate: this.ToDate,
      groups: this.groups,
      storeorgroup: this.storeorgrp,
      zeroro: this.zeroro,
      count: 0,
      Paytype: this.Paytype.toString(),
    };
    this.shared.api.SetHeaderData({
      obj: data,
    });
   

    this.getlaborData();
    // this.GetData('Rank', 'asc');
  }

  ngOnInit(): void {


    localStorage.setItem('time', this.dateType);

    // var curl = 'https://fbxtract.axelautomotive.com/favouritereports/GetServiceAdvisorRankingsV2';
    // this.apiSrvc.logSaving(curl, {}, '', 'Success', 'Service Advisor Rankings');
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
  //     if (this.storeorgrp == 'G' && (col_Name != 'Rank' && col_Name != 'ServiceAdvisor' && col_Name != 'StoreName')) {
  //       this.columnState = 'desc';
  //       this.columnName = col_Name;
  //       this.GetData(this.columnName, this.columnState);
  //     } else {
  //       this.columnState = 'asc';
  //       this.columnName = col_Name;
  //       this.GetData(this.columnName, this.columnState);
  //     }
  //   }
  //   //console.log(this.columnName, this.columnState);

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
  StoresData(data: any) {
  
    console.log(data, 'Data');

    this.storeIds = data.storeids;
    this.groupId = data.groupId;
    this.storename = data.storename;
    this.groupName = data.groupName;
    this.storecount = data.storecount;
    this.storedisplayname = data.storedisplayname;
  }
  getlaborData() {
    if (this.StoreVal != '' || this.selectedotherstoreids != '') {
      this.shared.spinner.show();
      if (this.LaborTypeVal == '') {
        const obj = {
          StoreID:
  this.storeIds.toString(),
          type: this.LaborState
        };
        this.shared.api.postmethod(this.comm.routeEndpoint + 'GetLaborTypesTechEfficiency', obj).subscribe((res: { response: any[]; }) => {
          this.labortypes = res.response
          this.selectedlabortypevalues = res.response.map(function (a: any) {
            return a.ASD_labortype;
          });
          // this.spinner.hide();
          // this.getlaborData();
          this.GetData('Rank', 'asc');

        })

      } else {
        this.shared.spinner.hide();
        this.GetData('Rank', 'asc');

      }
    } else {
      this.NoData = true
    }

  }

  GetData(sortdata?: any, sortstate?: any) {
    this.ServiceAdvisorData = [];
    this.shared.spinner.show();
    const obj = {
      StartDate: this.FromDate,
      EndDate: this.ToDate,
      StoreID:
       
 this.storeIds.toString(),

      Exp: sortdata,
      OrderType: sortstate,
      RankBy: this.storeorgrp,
      UserID: 0,
      LaborTypes: this.selectedlabortypevalues.toString(),
      ZeroHours: this.zeroro,
      Paytype: this.Paytype.toString(),
    };
    let startFrom = new Date().getTime();
    const curl = this.shared.getEnviUrl() + this.comm.routeEndpoint + 'GetServiceAdvisorRankingsV2';
    this.shared.api
      .postmethod(this.comm.routeEndpoint + 'GetServiceAdvisorRankingsV2', obj)
      .subscribe(
        (res: { message: any; status: number; response: string | any[] | undefined; }) => {
          const currentTitle = document.title;
          this.shared.api.logSaving(curl, {}, '', res.message, currentTitle);
          if (res.status == 200) {
            if (res.response != undefined) {
              if (res.response.length > 0) {
                let resTime = (new Date().getTime() - startFrom) / 1000;
                // this.logSaving(
                //   this.solutionurl + this.comm.routeEndpoint+'GetServiceAdvisorRankingsV2',
                //   obj,
                //   resTime,
                //   'Success'
                // );
                this.ServiceAdvisorData = res.response;
                // If Rank By Store → move TOTAL row to top

                if (this.storeorgrp.toString() === 'S') {
  const totalRowIndex = this.ServiceAdvisorData.findIndex(
    (x: any) => x.ServiceAdvisor === 'Total'
  );

  if (totalRowIndex > -1) {
    const totalRow = this.ServiceAdvisorData.splice(totalRowIndex, 1)[0];
    this.ServiceAdvisorData.unshift(totalRow);
  }
}


// // ---- KEEP TOTAL ROW IN CORRECT PLACE BASED ON RANK BY ----
// const totalIndex = this.ServiceAdvisorData.findIndex(
//   (x: any) => (x.ServiceAdvisor + '').toLowerCase() === 'total'
// );

// if (totalIndex > -1) {
//   const totalRow = this.ServiceAdvisorData.splice(totalIndex, 1)[0];

//   // Rank By Store (S) → TOP
//   if (this.storeorgrp.toString() === 'S') {
//     this.ServiceAdvisorData.unshift(totalRow);
//   }

//   // // Rank By Group (G) → BOTTOM
//   // if (this.storeorgrp.toString() === 'G') {
//   //   this.ServiceAdvisorData.push(totalRow);
//   // }
// }
                this.shared.spinner.hide();
                this.NoData = false;
                let position = this.scrollCurrentposition + 10
                setTimeout(() => {
                  this.scrollcent.nativeElement.scrollTop = position
                  //console.log(position);

                }, 500);
                // this.ServiceAdvisorData.some(function (x: any) {
                //   x.Data2 = JSON.parse(x.Data2);
                //   x.Dealerx = '+';
                //   return false;
                // });
                // this.GetTotalData();
              } else {
                // this.toast.error('Empty Response', '');
                this.shared.spinner.hide();
                this.NoData = true;
              }
            } else {
              let resTime = (new Date().getTime() - startFrom) / 1000;
              // this.logSaving(
              //   this.solutionurl + this.comm.routeEndpoint+'GetServiceAdvisorRankingsV2',
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
        (error: any) => {
          // this.toast.error('502 Bad Gate Way Error', '');
          this.shared.spinner.hide();
          this.NoData = true;
        }
      );
  }

  openDetails(data: any) {
    const DetailsServicePerson = this.shared.ngbmodal.open(
      // ServiceDetailsV3Component,
      {
        // size:'xl',
        backdrop: 'static',
      }
    );
    DetailsServicePerson.componentInstance.Servicedetails = [
      {
        StartDate: this.FromDate,
        EndDate: this.ToDate,
        var1: 'Store_Name',
        var2: 'ServiceAdvisor_Name',
        var3: '',
        var1Value: data.StoreName,
        var2Value: data.ServiceAdvisor,
        var3Value: '',
        PaytypeC: 'C',
        PaytypeW: 'W',
        PaytypeI: 'I',
        DepartmentS: 'S',
        DepartmentP: 'P',            // DepartmentP: '',
        DepartmentQ: 'Q',
        DepartmentB: 'B',
        PolicyAccount: 'N',
        userName: data.ServiceAdvisor,
        Grosstype: '',
        layer: 2,
        zeroHours: this.zeroro == 'I' ? 'Y' : ''
      },
    ];
  }

  reportOpen(temp: any) {


    // this.ngbmodalactive = this.ngbmodal.open(temp, {
    //   size: 'xl',
    //   backdrop: 'static',
    // });
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

      var arr = this.ServiceAdvisorData.slice(
        1,
        this.ServiceAdvisorData.length
      );
      arr.sort(function (a: any, b: any) {
        if (a[property] < b[property]) {
          return -1 * direction;
        } else if (a[property] > b[property]) {
          return 1 * direction;
        } else {
          return 0;
        }
      });
      arr.unshift(this.ServiceAdvisorData[0]);
      this.ServiceAdvisorData = arr;   

//       // KEEP TOTAL IN CORRECT PLACE AFTER SORT
// const totalIndex = arr.findIndex(
//   (x: any) => (x.ServiceAdvisor + '').toLowerCase() === 'total'
// );

// if (totalIndex > -1) {
//   const totalRow = arr.splice(totalIndex, 1)[0];

//   if (this.storeorgrp.toString() === 'S') {
//     arr.unshift(totalRow); // store → top
//   } else {
//     arr.push(totalRow); // group → bottom
//   }
// }
  }
  SARstate: any;
  ngAfterViewInit(): void {
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.comm.pageName == 'Service Advisor Rankings') {
        if (res.obj.storesData != undefined) {
          this.groupsArray = res.obj.storesData;
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
    this.reportOpenSub = this.shared.api.GetReportOpening().subscribe((res: { obj: { Module: string; }; }) => {
      //console.log(res);
      if (this.reportOpenSub != undefined) {
        if (res.obj.Module == 'Service Advisor Rankings') {
          document.getElementById('report')?.click()
        }
      }
    });
    this.reportGetting = this.shared.api.GetReports().subscribe((data: { obj: { Reference: string; header: string | undefined; PayType: any; labortypeValues: any; TotalReport: any; storeorgroup: any; zeroro: any; laborstate: any; otherstoreids: any; groups: any; FromDate: undefined; ToDate: undefined; storeValues: any; dateType: any; }; }) => {
      if (this.reportGetting != undefined) {

        if (data.obj.Reference == 'Service Advisor Rankings') {
          if (data.obj.header == undefined) {
            this.Paytype = data.obj.PayType
            // this.StoreVal = data.obj.storeValues;
            this.LaborTypeVal = data.obj.labortypeValues;         
            this.storeorgrp = data.obj.storeorgroup;
            this.zeroro = data.obj.zeroro;
            this.LaborState = data.obj.laborstate;
            this.selectedotherstoreids = data.obj.otherstoreids;
            this.groups = data.obj.groups
            if (data.obj.FromDate != undefined && data.obj.ToDate != undefined) {
              this.FromDate = data.obj.FromDate;
              this.ToDate = data.obj.ToDate;
              this.StoreVal = data.obj.storeValues;
              this.dateType = data.obj.dateType;
              this.LaborState = data.obj.laborstate;
              this.selectedotherstoreids = data.obj.otherstoreids;
              this.getlaborData();
              // this.GetData(this.columnName, this.columnState);
            } else {
              this.FromDate = data.obj.FromDate;
              this.ToDate = data.obj.ToDate;
              this.StoreVal = data.obj.storeValues;
              this.dateType = data.obj.dateType;
              this.LaborState = data.obj.laborstate;
              this.selectedotherstoreids = data.obj.otherstoreids;
              this.getlaborData();
              // this.GetData(this.columnName, this.columnState);
            }
          }
          else {
            if (data.obj.header == 'Yes') {
              // this.storeIds = data.obj.storeValues;
              this.StoreVal = data.obj.storeValues
              //console.log(this.storeIds);
              this.getlaborData();
              // this.GetData(this.columnName, this.columnState);

            }
          }
          const headerdata = {
            title: 'Service Advisor Rankings',
            path1: '',
            path2: '',
            path3: '',
            stores:  this.storeIds.toString(),
            labortype: this.LaborTypeVal.toString(),
          
            datetype: this.dateType,
            fromdate: this.FromDate,
            todate: this.ToDate,
            groups: this.groups,
            storeorgroup: this.storeorgrp,
            zeroro: this.zeroro,
            Paytype: this.Paytype,
            // labortype:this.LaborTypeVal.toString(),
            laborstate: this.LaborState
          };
          this.shared.api.SetHeaderData({
            obj: headerdata,
          });
        
        }
      }
    });

 

    this.excel = this.shared.api.getExportToExcelAllReports().subscribe((res: { obj: { state: boolean; title: string; }; }) => {
      if (this.excel != undefined) {
        this.SARstate = res.obj.state;
        if (res.obj.title == 'Service Advisor Rankings') {
          if (res.obj.state == true) {
            this.exportToExcel();
          }
        }
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res: { obj: { title: string; statePrint: boolean; }; }) => {
      if (this.print != undefined) {
        if (res.obj.title == 'Service Advisor Rankings') {
          if (res.obj.statePrint == true) {
            // this.GetPrintData();
          }
        }
      }
    });

    this.Pdf = this.shared.api.getExportToPDFAllReports().subscribe((res: { obj: { title: string; statePDF: boolean; }; }) => {
      if (this.Pdf != undefined) {
        if (res.obj.title == 'Service Advisor Rankings') {
          if (res.obj.statePDF == true) {
            // this.generatePDF();
          }
        }
      }
    });
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res: { obj: { title: string; stateEmailPdf: boolean; }; }) => {
      if (this.email != undefined) {
        if (res.obj.title == 'Service Advisor Rankings') {
          if (res.obj.stateEmailPdf == true) {
            // this.sendEmailData(res.obj.Email, res.obj.notes, res.obj.from);
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
  // currentElement: string;

  // @ViewChild('scrollOne') scrollOne: ElementRef;
  // @ViewChild('scrollTwo') scrollTwo: ElementRef;

  // updateVerticalScroll(event): void {
  //   if (this.currentElement === 'scrollTwo') {
  //     this.scrollOne.nativeElement.scrollTop = event.target.scrollTop;
  //   } else if (this.currentElement === 'scrollOne') {
  //     this.scrollTwo.nativeElement.scrollTop = event.target.scrollTop;
  //   }
  // }

  // updateCurrentElement(element: 'scrollOne' | 'scrollTwo') {
  //   this.currentElement = element;
  // }

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



  //------Reports-----------//
  activePopover: number = -1;
  Teams = 'S';
  selectedstorevalues: any = [];
  labortypes: any = [];
  selectedlabortypevalues: any = [];
  Performance: any = 'Load';
  changesvalues: any = [];
  AllLabortypes: boolean = true;
  toporbottom: any = ['T'];

  initializeDates(type: any) {
    let dates: any = this.setdates.setDates(type)
    this.FromDate = dates[0];
    this.ToDate = dates[1];
    localStorage.setItem('time', type);
    this.setDates(type)
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
  teamsselection(e: any) {
    this.Teams = e;
    this.getlabortype(2, e)
  }
  spinnerLoaderlabor: boolean = false;
  getlabortype(count?: any, block?: any, popuporbar?: any) {
    // if(this.selectedstorevalues.length==this.stores.length){
    //   this.selectedstorevalues=''
    // }
    console.log(count, block, popuporbar);
    let allstrids: any = [];
    allstrids = [...this.selectedotherstoreids, ...this.selectedstorevalues]
    this.spinnerLoaderlabor = true;
    this.labortypes = [];
    const obj = {
      StoreID: allstrids,
      type: block
    };
    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetLaborTypesTechEfficiency', obj).subscribe((res: { response: any; }) => {
      this.spinnerLoaderlabor = false;
      this.labortypes = res.response;
      this.selectedlabortypevalues = []
      //console.log(this.selectedlabortypevalues.length, this.labortypes.length,count);

      this.shared.api.GetHeaderData().subscribe((res: { obj: { title: string; labortype: string; }; }) => {
        //console.log(res);

        if (
          res.obj.title == 'Service Advisor Rankings' &&
          this.Performance == 'Load' &&
          // this.groupstate == false &&
          count != 2 && popuporbar != 'Bar'
        ) {
          if (res.obj.labortype != '' && res.obj.labortype != '0') {
            this.selectedlabortypevalues = [];
            if (res.obj.labortype.toString().indexOf(',') > 0) {
              var result = res.obj.labortype.split(',');
              //console.log(result);

              this.selectedlabortypevalues = result;
              //console.log(this.selectedlabortypevalues);

              // .map(function (x: any) {
              //   return parseInt(x, 10);
              // });
            } else {
              this.selectedlabortypevalues.push(res.obj.labortype);
              //console.log(this.selectedlabortypevalues);

            }
          }
          if (res.obj.labortype == '' || res.obj.labortype == '0') {
            //console.log(this.labortypes);
            this.selectedlabortypevalues = [];

            this.selectedlabortypevalues = this.labortypes.map(function (
              a: any
            ) {
              return a.ASD_labortype;
            });
            //console.log(this.selectedlabortypevalues);
          }
        }


      });
      if (popuporbar == 'Bar') {
        if (this.changesvalues.labortype != '' && this.changesvalues.labortype != '0') {
          this.selectedlabortypevalues = [];
          if (this.changesvalues.labortype.toString().indexOf(',') > 0) {
            var result = this.changesvalues.labortype.split(',');
            this.selectedlabortypevalues = result;
            // .map(function (x: any) {
            //   return parseInt(x, 10);
            // });
            //console.log(this.selectedlabortypevalues);

          } else {
            this.selectedlabortypevalues.push(this.changesvalues.labortype);
          }
        }
        if (this.changesvalues.labortype == '' || this.changesvalues.labortype == '0') {
          //console.log(this.labortypes);
          this.selectedlabortypevalues = [];

          this.selectedlabortypevalues = this.labortypes.map(function (
            a: any
          ) {
            return a.ASD_labortype;
          });
          //console.log(this.selectedlabortypevalues);
        }
      }
      else {
        this.selectedlabortypevalues = this.labortypes.map(function (a: any) {
          return a.ASD_labortype;
        });
      }
    });
  }
  alllabortypes(type: any) {
    this.AllLabortypes = !this.AllLabortypes;

    if (type == 'Y') {
      this.selectedlabortypevalues = this.labortypes.map(function (a: any) {
        return a.ASD_labortype;
      });
    }
    else {
      this.selectedlabortypevalues = [];
    }
  }
  individualLabortypes(e: any) {
    console.log(e);

    const index = this.selectedlabortypevalues.findIndex(
      (i: any) => i == e.ASD_labortype
    );
    if (index >= 0) {
      this.selectedlabortypevalues.splice(index, 1);
      this.AllLabortypes = false;
    } else {
      this.selectedlabortypevalues.push(e.ASD_labortype);
      if (this.selectedlabortypevalues.length == this.labortypes.length) {
        this.AllLabortypes = true;
      } else {
        this.AllLabortypes = false;
      }
    }
  }
  zeroros(block: any, e: any) {
    // if (block == 'TB') {
    this.zeroro = [];
    this.zeroro.push(e);
    // }
  
  }
  storeorgroups(block: any, e: any) {
    // if (block == 'TB') {
    this.storeorgrp = [];
    this.storeorgrp.push(e);
    // }
  }
  multipleorsingle(block: any, e: any) {
    if (block == 'TB') {
      this.toporbottom = [];
      this.toporbottom.push(e)
    }

    if (block == 'PT') {

      const index = this.Paytype.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.Paytype.splice(index, 1);
      } else {
        this.Paytype.push(e);
      }
    }
  }
  viewreport() {
    this.activePopover = -1
    // if (this.selectedlabortypevalues.length == this.labortypes.length) {
    //   this.selectedlabortypevalues = 0;
    // }
//    const hasStoreSelection =
//   (Array.isArray(this.selectedstorevalues) && this.selectedstorevalues.length > 0) ||
//   (typeof this.selectedstorevalues === 'string' && this.selectedstorevalues.trim() !== '') ||
//   (Array.isArray(this.selectedotherstoreids) && this.selectedotherstoreids.length > 0);

// if (!hasStoreSelection) {

//   return;
// }

  if(!this.storeIds || this.storeIds.length === 0){

      this.toast.show('Please Select Atleast One Store', 'warning', 'Warning');
    }
 else 
  
  if (!this.selectedlabortypevalues || this.selectedlabortypevalues.length === 0) {
  
    this.toast.show('Please select any labor type', 'warning', 'Warning');
    return;
  }

 else if (
    !this.Paytype ||
    (Array.isArray(this.Paytype) && this.Paytype.length === 0)
  ) {
  
    this.toast.show('Please select any Pay Type', 'warning', 'Warning');
    return;
  }
    else {
      const data = {
        Reference: 'Service Advisor Rankings',
        FromDate: this.FromDate,
        ToDate: this.ToDate,
        TotalReport: this.toporbottom[0],
        storeValues: this.storeIds.toString(),
        labortypeValues:
          this.selectedlabortypevalues.toString(),
        //  == '0'
        //   ? ''
        //   : this.selectedlabortypevalues.toString(),
        laborstate: this.Teams,
        dateType: this.DateType,
        // groups: this.selectedGroups.toString(),
        storeorgroup: this.storeorgrp.toString(),
        zeroro: this.zeroro.toString(),
        otherstoreids: this.selectedotherstoreids,
        PayType: this.Paytype,

      };
      // this.close();

      this.shared.api.SetReports({
        obj: data,
      });
          this.GetData(this.columnName, this.columnState);

    }

  }
  ExcelStoreNames: any = []
  exportToExcel(): void {
    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('Service Advisor Rankings');
    const title = worksheet.addRow(['Service Advisor Rankings']);
    title.font = { size: 14, bold: true, name: 'Arial' };
    title.alignment = { vertical: 'middle', horizontal: 'center' };
    worksheet.mergeCells('A1:Q1');
    worksheet.addRow([]);
    const formattedFromDate = this.shared.datePipe.transform(this.FromDate, 'dd-MMM-yyyy');
    const formattedToDate = this.shared.datePipe.transform(this.ToDate, 'dd-MMM-yyyy');
  
     const value = ((this.storeorgrp !== undefined && this.storeorgrp !== null && this.storeorgrp.toString().length > 0)
    ? this.storeorgrp.toString()
    : (this.storeorgrp ? this.storeorgrp.toString() : '')).toUpperCase();

  let rankByValue = 'Store';
  if (value === 'G' || value === 'GROUP') {
    rankByValue = 'Group';
  } else if (value === 'S' || value === 'STORE') {
    rankByValue = 'Store';
  }

      const payTypeMap: any = {
        'C': 'Customer Pay',
        'W': 'Warranty',
        'I': 'Internal',
        'S': 'Sublet Gross',
        'M': 'Misc',
      };
      
      // Convert Paytype array or single value to readable names
      let formattedPayType = 'All';
      if (Array.isArray(this.Paytype) && this.Paytype.length > 0) {
        formattedPayType = this.Paytype.map(pt => payTypeMap[pt] || pt).join(', ');
      } else if (typeof this.Paytype === 'string' && this.Paytype !== '') {
        formattedPayType = payTypeMap[this.Paytype] || this.Paytype;
      }
      
      // Handle Zero Hours value
      let formattedZeroHours = 'All';
      if (this.zeroro) {
        formattedZeroHours = this.zeroro === 'E' ? 'Exclude' :
                             this.zeroro === 'I' ? 'Include' :
                             this.zeroro;
      }
      
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
        { name: 'Labor Types:', values: this.selectedlabortypevalues },
        { name: 'Zero Hours:', values: formattedZeroHours },
        { name: 'Rank By:', values: this.storeorgrp == 'S' ? 'Store' :'Group' },
        { name: 'Pay Type:', values: formattedPayType },
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

  
  
    const firstHeader = ['', '', '', '', '', '','','', '', '', '','','','','','',''];
    const headerRow1 = worksheet.addRow(firstHeader);
  

    const secondHeader = [
      'Rank', '	Service Advisor', 'Store Name','#ROs	', 'ROs/Day	', 'Hours	', 'Hours/Ro','	Labor Sales','Labor Gross', 'Parts Sales	', 'Parts Gross', 'Total Gross','Sublet Gross','Misc','GP%','Discounts','ELR'
    ];
    const headerRow2 = worksheet.addRow(secondHeader.map(h => h.trim()));
  
    const headerRow1Index = headerRow1.number;
    const headerRow2Index = headerRow2.number;
  
   
    worksheet.mergeCells(`A${headerRow1Index}`);
    worksheet.mergeCells(`B${headerRow1Index}`);
    worksheet.mergeCells(`C${headerRow1Index}`);
    worksheet.mergeCells(`D${headerRow1Index}:H${headerRow1Index}`); // Back Gross
    worksheet.mergeCells(`I${headerRow1Index}:L${headerRow1Index}`); // Unit Credit
  
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
      'Rank', 'ServiceAdvisor', 'StoreName',
      'RO', 'ROPerDay', 
     'ActualHours','HoursPerRO','LaborSale' ,'LaborGross','PartsSale' ,'PartsGross','TotalGross','SubletGross','MiscGross','GP','Discount','ELR'
    ];
  
    const currencyFields = ['LaborSale' ,'LaborGross','PartsSale' ,'PartsGross','TotalGross','ELR'];
  
    this.ServiceAdvisorData.forEach((info: any) => {
    const rowData = bindingHeaders.map((key) => {
  let val = info[key];

  // Suppress rank 1000 only for Total row
  if (key === 'Rank' && (info.ServiceAdvisor + '').toLowerCase() === 'total') {
    return '';   // leave blank
  }

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
      this.shared.exportToExcel(workbook, 'Service Advisor Rankings');
    });
  }
}
