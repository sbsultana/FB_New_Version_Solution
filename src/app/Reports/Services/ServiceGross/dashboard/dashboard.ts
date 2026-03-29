import { Component, OnInit, ViewChild, ElementRef, HostListener } from '@angular/core';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { Setdates } from '../../../../Core/Providers/SetDates/setdates';
import { BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { common } from '../../../../common';
import { DateRangePicker } from '../../../../CommonFilters/date-range-picker/date-range-picker';
import { Subscription } from 'rxjs';
import { environment } from '../../../../../environments/environment';
import { Stores } from '../../../../CommonFilters/stores/stores';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
import { CurrencyPipe } from '@angular/common';
import { NgbModalRef } from '@ng-bootstrap/ng-bootstrap';
import { ServicegrossDetails } from '../servicegross-details/servicegross-details';

@Component({
  selector: 'app-dashboard',
  imports: [SharedModule, BsDatepickerModule, DateRangePicker, Stores],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss'
})
export class Dashboard {
  ServiceData: any = [];
  IndividualServiceGross: any = [];
  TotalServiceGross: any = [];

  IndividualServicePaytype: any = [];
  TotalServicePaytype: any = [];
  ServicePaytype: any = [];

  IndividualServiceHour: any = [];
  TotalServiceHour: any = [];
  ServiceHour: any = [];

  IndividualServiceRO: any = [];
  TotalServiceRO: any = [];
  ServiceRO: any = [];

  type: any = 'A';
  TotalReport: string = 'T';
  NoData: boolean = false;
  ROType: any = 'Closed';
  Department: any = ['Service'];
  Paytype: any = ['Customerpay_0', 'Warranty_1', 'Internal_2', 'Extended_3'];
  Grosstype: any = ['Labour_0', '', 'Misc_2', 'Sublet_3'];
  // zeroHours: any = ['N']
  labourType: any = 'N';
  GridView = 'Global';
  Target: any = 'F';

  responcestatus: any = '';
  CurrentDate = new Date();
  spinnerLoader: boolean = false;
  getState: any = '';
  zeroHours: any = true;
  Sublet: any = false;
  Misc: any = false;
  Lube: any = false
  actionType: any = ''
  DefaultLoad: any = 'E'

  reportOpenSub!: Subscription;
  reportGetting!: Subscription;
  Pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;

  LaborTypeVal: any = []

  stores: any = []
  groupsArray: any = [];
  storename: any = ''
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  groupId: any = 0;
  storeIds: any = 0;
  otherStoresArray: any = [];
  otherStoreIds: any = [];


  storesFilterData: any = {
    'groupsArray': this.groupsArray, 'groupId': this.groupId, 'storesArray': this.stores, 'storeids': '1', 'type': 'M', 'others': 'Y', otherStoresArray: this.otherStoresArray, otherStoreIds: this.otherStoreIds,
    'groupName': this.groupName, 'storename': this.storename, storecount: null, 'storedisplayname': this.storedisplayname, 'DefaultLoad': this.DefaultLoad
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



  selectedDataGrouping: any = [];
  dataGrouping: any = [
    { "ARG_ID": 40, "ARG_LABEL": "Store", "ARG_SEQ": 0, "state": true, "columnname": "Store_Name", "Active": "Y" },
    { "ARG_ID": 26, "ARG_LABEL": "Advisor Name", "ARG_SEQ": 1, "state": true, "Active": "Y", "columnname": "ServiceAdvisor_Name" },
    { "ARG_ID": 27, "ARG_LABEL": "Advisor Number", "ARG_SEQ": 2, "state": false, "Active": "Y", "columnname": "ServiceAdvisor" },
    { "ARG_ID": 28, "ARG_LABEL": "Tech Name", "ARG_SEQ": 3, "state": false, "Active": "Y", "columnname": "techname" },
    { "ARG_ID": 29, "ARG_LABEL": "Tech Number", "ARG_SEQ": 4, "state": false, "Active": "Y", "columnname": "techno" },
    { "ARG_ID": 30, "ARG_LABEL": "Pay Type", "ARG_SEQ": 5, "state": false, "Active": "N", "columnname": "" },
    { "ARG_ID": 31, "ARG_LABEL": "Vehicle Year", "ARG_SEQ": 6, "state": false, "Active": "Y", "columnname": "vehicle_Year" },
    { "ARG_ID": 32, "ARG_LABEL": "Vehicle Make", "ARG_SEQ": 7, "state": false, "Active": "Y", "columnname": "vehicle_Make" },
    { "ARG_ID": 33, "ARG_LABEL": "Vehicle Model", "ARG_SEQ": 8, "state": false, "Active": "Y", "columnname": "Vehicle_Model" },
    { "ARG_ID": 34, "ARG_LABEL": "Vehicle Odometer", "ARG_SEQ": 9, "state": false, "Active": "Y", "columnname": "Vehicle_Odometer" },
    { "ARG_ID": 35, "ARG_LABEL": "Customer Name", "ARG_SEQ": 10, "state": false, "Active": "Y", "columnname": "CName" },
    { "ARG_ID": 36, "ARG_LABEL": "Customer ZIP", "ARG_SEQ": 11, "state": false, "Active": "Y", "columnname": "CZip" },
    { "ARG_ID": 37, "ARG_LABEL": "Customer State", "ARG_SEQ": 12, "state": false, "Active": "Y", "columnname": "CState" },
    { "ARG_ID": 38, "ARG_LABEL": "RO Close Date", "ARG_SEQ": 13, "state": false, "Active": "Y", "columnname": "cdate" },
    { "ARG_ID": 39, "ARG_LABEL": "No Grouping", "ARG_SEQ": 14, "state": false, "Active": "N", "columnname": "" },
    { "ARG_ID": 68, "ARG_LABEL": "Opcode", "ARG_SEQ": 15, "state": false, "Active": "Y", "columnname": "opcode" }
  ];

  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent) {
    const clickedInside = (event.target as HTMLElement).closest('.dropdown-toggle, .reportstores-card, .timeframe');
    if (!clickedInside) {
      this.activePopover = -1;
    }
  }
  constructor(
    public shared: Sharedservice, public setdates: Setdates, private comm: common, private cp: CurrencyPipe, private toast: ToastService,
  ) {

    this.selectedDataGrouping.push(this.dataGrouping[0])
    this.selectedDataGrouping.push(this.dataGrouping[1])

    this.shared.setTitle(this.comm.titleName + '-Service Gross');

    if (localStorage.getItem('flag') == 'V') {
      this.storeIds = [];
      console.log(JSON.parse(localStorage.getItem('userInfo')!), JSON.parse(localStorage.getItem('userInfo')!).user_Info, 'Widget Stores............');
      this.groupId = JSON.parse(localStorage.getItem('userInfo')!).groupid
      JSON.parse(localStorage.getItem('userInfo')!).store.indexOf(',') > 0 ?
        this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).store.split(',') :
        this.storeIds.push(JSON.parse(localStorage.getItem('userInfo')!).store)
      this.actionType = 'Y';
      this.Department = ['Service'];
      this.zeroHours = false;
      localStorage.setItem('flag', 'M')
    } else {
      if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
        this.groupId = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Preferences
        this.storeIds = ''
        this.otherStoreIds = [];

        this.actionType = 'N';
        this.Department = ['Service', 'Parts', 'Quicklube']
        this.zeroHours = true;
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
    this.headervalues();
    this.getlabourTypesData('FL', 'N');

  }
  scrollPosition = 0;
  getScrollPosition(event: any): void {
    this.scrollPosition = event.target.scrollLeft;
  }
  labortypes: any = []
  getlabourTypesData(block: any, type: any) {
    if (this.storeIds != '') {
      this.spinnerLoaderlabor = true;
      this.responcestatus = '';

      // this.shared.spinner.show();    
      const obj = {
        StoreID: [...this.storeIds, ...this.otherStoreIds],
        type: type == 'N' ? 'A' : type
      };
      this.spinnerLoader = true;
      this.shared.api.postmethod(this.comm.routeEndpoint + 'GetLaborTypesTechEfficiency', obj).subscribe((res) => {
        this.spinnerLoaderlabor = false;
        this.spinnerLoader = false;
        this.labortypes = res.response;
        this.LaborTypeVal = res.response.map(function (a: any) {
          return a.ASD_labortype;
        });
        if (block == 'FL') {
          this.getServiceData()
        }
      })
    } else {
      // this.NoData = true
    }
  }

  initializeDates(type: any) {
    let dates: any = this.setdates.setDates(type)
    this.FromDate = dates[0];
    this.ToDate = dates[1];
    localStorage.setItem('time', type);
    this.DateType = type
    this.setDates(this.DateType)

  }

  headervalues() {
    const data = {
      title: 'Service Gross',
      dataGroupings: this.selectedDataGrouping,
      ROType: this.ROType,
      Department: this.Department,
      Paytype: this.Paytype,
      Target: this.Target,
      ToporBottom: this.TotalReport,
      fromdate: this.FromDate,
      todate: this.ToDate,
      grosstype: this.Grosstype,
      GridView: this.GridView,
      groups: this.groupId,
      labourtype: this.labourType,
      zeroHours: this.zeroHours,
      labourtypevalues: this.LaborTypeVal,
      otherstoreids: this.otherStoreIds
    };
    this.shared.api.SetHeaderData({ obj: data });


  }
  getServiceData() {
    if (this.storeIds != '' || this.otherStoreIds != '') {
      this.actionType = 'Y'
      this.DefaultLoad = ''
      if (this.getState.indexOf('S') < 0) {
        this.responcestatus = '';
        this.shared.spinner.show();
        this.GetData();
        this.GetTotalData();
        this.getState = this.getState + 'S'
      }
    } else {
      this.NoData = true
    }
  }

  viewRO(roData: any) {

  }
  GetData() {
    this.IndividualServiceGross = [];
    const obj = {
      startdate: this.FromDate,
      enddate: this.ToDate,
      StoreID: [...this.storeIds, ...this.otherStoreIds],
      // StoreID:'27,14,15,3,23,7,28,13,29,11,10,30,16,17,8,1,26,20,21,6,25,18,9,4,19',

      AdvisorNumber: '',
      AdvisorName: '',
      ROSTATUS: '',
      PaytypeC: this.Paytype[0] == 'Customerpay_0' ? 'C' : '',
      PaytypeW: this.Paytype[1] == 'Warranty_1' ? 'W' : '',
      PaytypeI: this.Paytype[2] == 'Internal_2' ? 'I' : '',
      PaytypeE: this.Paytype[3] == 'Extended_3' ? 'E' : '',
      DepartmentS: this.Department.indexOf('Service') >= 0 ? 'S' : '',
      DepartmentP: this.Department.indexOf('Parts') >= 0 ? 'P' : '',
      // DepartmentP: '',
      DepartmentQ: this.Department.indexOf('Quicklube') >= 0 ? 'Q' : '',
      DepartmentB: this.Department.indexOf('Body') >= 0 ? 'B' : '',
      DepartmentD: this.Department.indexOf('Details') >= 0 ? 'D' : '',
      GrossTypeM: this.Misc == true ? 'M' : '',
      GrossTypeL: this.Lube == true ? 'L' : '',
      GrossTypeS: this.Sublet == true ? 'S' : '',
      GrossTypeP: '',
      PolicyAccount: this.labourType == 'Y' ? this.labourType : 'N',
      excludeZeroHours: this.zeroHours == true ? 'Y' : 'N',
      vehicle_Year: '',
      vehicle_Make: '',
      Vehicle_Model: '',
      Vehicle_Odometer: '',
      CName: '',
      CZip: '',
      CState: '',
      RO_CloseDate: '',
      var1: this.selectedDataGrouping.length >= 1 ? this.selectedDataGrouping[0]?.columnname : '',
      var2: this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '',
      var3: this.selectedDataGrouping.length == 3 ? this.selectedDataGrouping[2]?.columnname : '',
      type: 'D',
      LaborTypes: this.LaborTypeVal.toString()
    };
    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetServiceSummaryBetaDetailsV2', obj).subscribe(
      (res) => {
        if (res.status == 200) {
          if (res.response != undefined) {
            if (res.response.length > 0) {
              this.IndividualServiceGross = res.response;
              this.responcestatus = this.responcestatus + 'I';
              this.NoData = false;
              let path2 = this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '';
              this.IndividualServiceGross.some(function (x: any) {
                if (x.Data2 != undefined) {
                  x.Data2 = JSON.parse(x.Data2);
                  x.Data2 = x.Data2.map((v: any) => ({
                    ...v,
                    SubData: [],
                    data2sign: '-',
                  }));
                }
                if (x.Data3 != undefined) {
                  x.Data3 = JSON.parse(x.Data3);
                  x.Data2.forEach((val: any) => {
                    x.Data3.forEach((ele: any) => {
                      if (val.data2 == ele.data2) {
                        val.SubData.push(ele);
                      }
                    });
                  });
                }
                if (path2 == '') {
                  x.Dealer = '+';
                } else {
                  x.Dealer = '-';
                }
              });
              this.combineIndividualandTotal();
            } else {
              // this.toast.show('Empty Response','');
              this.shared.spinner.hide();
              this.NoData = true;
            }
          } else {
            // this.toast.show('Empty Response','');
            this.shared.spinner.hide();
            this.NoData = true;
          }
        } else {
          //this.toast.show('No Response', '');
          this.shared.spinner.hide();
          this.NoData = true;
        }
      },
      (error) => {
        this.toast.show('502 Bad Gate Way Error', 'danger', 'Error');
        this.shared.spinner.hide();
        this.NoData = true;
      }
    );
  }
  GetTotalData() {
    this.TotalServiceGross = [];
    const obj = {
      startdate: this.FromDate,
      enddate: this.ToDate,
      StoreID: [...this.storeIds, ...this.otherStoreIds],
      AdvisorNumber: '',
      AdvisorName: '',
      ROSTATUS: '',
      PaytypeC: this.Paytype[0] == 'Customerpay_0' ? 'C' : '',
      PaytypeW: this.Paytype[1] == 'Warranty_1' ? 'W' : '',
      PaytypeI: this.Paytype[2] == 'Internal_2' ? 'I' : '',
      PaytypeE: this.Paytype[3] == 'Extended_3' ? 'E' : '',

      DepartmentS: this.Department.indexOf('Service') >= 0 ? 'S' : '',
      DepartmentP: this.Department.indexOf('Parts') >= 0 ? 'P' : '',
      // DepartmentP: '',

      DepartmentQ: this.Department.indexOf('Quicklube') >= 0 ? 'Q' : '',
      DepartmentB: this.Department.indexOf('Body') >= 0 ? 'B' : '',
      DepartmentD: this.Department.indexOf('Details') >= 0 ? 'D' : '',
      GrossTypeM: this.Misc == true ? 'M' : '',
      GrossTypeL: this.Lube == true ? 'L' : '',
      GrossTypeS: this.Sublet == true ? 'S' : '',
      GrossTypeP: '',
      PolicyAccount: this.labourType == 'Y' ? this.labourType : 'N',
      excludeZeroHours: this.zeroHours == true ? 'Y' : 'N',

      vehicle_Year: '',
      vehicle_Make: '',
      Vehicle_Model: '',
      Vehicle_Odometer: '',
      CName: '',
      CZip: '',
      CState: '',
      RO_CloseDate: '',
      var1: this.selectedDataGrouping[0].columnname == 'Store_Name' ? '' : this.selectedDataGrouping[0].columnname,
      var2: this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '',
      var3: this.selectedDataGrouping.length == 3 ? this.selectedDataGrouping[2]?.columnname : '',
      type: 'T',
      LaborTypes: this.LaborTypeVal.toString()

    };
    const curl = environment.apiUrl + this.comm.routeEndpoint + 'GetServiceSummaryBetaDetailsV2';
    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetServiceSummaryBetaDetailsV2', obj).subscribe(
      (totalres) => {
        const currentTitle = document.title;
        this.shared.api.logSaving(curl, {}, '', totalres.message, currentTitle);
        if (totalres.status == 200) {
          if (totalres.response != undefined) {
            if (totalres.response.length > 0) {
              this.TotalServiceGross = totalres.response.map((v: any) => ({
                ...v,
                data1: 'REPORTS TOTAL',
                Dealer: '+',
                Data2: [],
              }));
              this.responcestatus = this.responcestatus + 'T';
              this.combineIndividualandTotal();
            } else {
              // this.toast.show('Empty Response','');
              this.shared.spinner.hide();
              this.NoData = true;
            }
          } else {
            // this.toast.show('Empty Response','');
            this.shared.spinner.hide();
            this.NoData = true;
          }
        } else {
          //this.toast.show('No Response', '');
          this.shared.spinner.hide();
          this.NoData = true;
        }
      },
      (error) => {
        this.toast.show('502 Bad Gate Way Error', 'danger', 'Error');
        this.shared.spinner.hide();
        this.NoData = true;
      }
    );
  }

  combineIndividualandTotal() {
    if (this.responcestatus == 'IT' || this.responcestatus == 'TI') {
      if (this.TotalReport == 'B') {
        this.IndividualServiceGross.push(this.TotalServiceGross[0]);
        this.ServiceData = this.IndividualServiceGross;
      } else {
        this.IndividualServiceGross.unshift(this.TotalServiceGross[0]);
        this.ServiceData = this.IndividualServiceGross;
      }
      this.shared.spinner.hide();
    } else if (this.responcestatus == 'T') {
      this.ServiceData = this.TotalServiceGross;
    } else if (this.responcestatus == 'I') {
      this.ServiceData = this.IndividualServiceGross;
    } else {
      this.NoData = true;
    }
  }

  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    }
    else if (value < 0) {
      return false;
    }
    return true
  }

  public getColorClass(value: number | null | undefined): string {
    if (value === null || value === undefined || value === 0) {
      return ''; // No class applied
    }
    return value > 0 ? 'positivebg' : 'negativebg';
  }

  expandorcollapse(ind: any, e: any, ref: any, Item: any, parentData: any) {
    let id = (e.target as Element).id;
    if (id == 'D_' + ind) {
      if ((this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '') == '') {
        this.openDetails(Item, parentData, '1');
      } else {
        if (ref == '-') {
          Item.Dealer = '+';
        }
        if (ref == '+') {
          Item.Dealer = '-';
        }
      }
    }
    if (id == 'DN_' + ind) {

      if ((this.selectedDataGrouping.length == 3 ? this.selectedDataGrouping[2]?.columnname : '') == '') {
        this.openDetails(Item, parentData, '2');
      } else {
        if (ref == '-') {
          Item.data2sign = '+';
        }
        if (ref == '+') {
          Item.data2sign = '-';
          Item.Dealer = '-';
        }
      }
    }
  }
  isDesc: boolean = false;
  column: string = 'CategoryName';

  sort(property: any, data: any) {
    this.isDesc = !this.isDesc; //change the direction
    this.column = property;
    let direction = this.isDesc ? 1 : -1;
    // //console.log(property)
    data.sort(function (a: any, b: any) {
      if (a[property] < b[property]) {
        return -1 * direction;
      } else if (a[property] > b[property]) {
        return 1 * direction;
      } else {
        return 0;
      }
    });
  }
  openDetails(Item: any, ParentItem: any, ref: any) {
    // //console.log(Item, ParentItem);
    if (ref == '1') {
      if (Item.data1 != undefined && Item.data1 != 'Reports Total') {
        const DetailsServicePerson = this.shared.ngbmodal.open(
          ServicegrossDetails,
          {
            // size:'xl',
            backdrop: 'static',
          }
        );
        DetailsServicePerson.componentInstance.Servicedetails = [
          {
            // "storeId": Item.data1,
            // "SrvcName": Item.ServiceAdvisor_Name,
            StartDate: this.FromDate,
            EndDate: this.ToDate,
            var1: this.selectedDataGrouping[0]?.columnname,
            var2: '',
            var3: '',
            var1Value: Item.data1,
            var2Value: '',
            var3Value: '',
            PaytypeC: this.Paytype[0] == 'Customerpay_0' ? 'C' : '',
            PaytypeW: this.Paytype[1] == 'Warranty_1' ? 'W' : '',
            PaytypeI: this.Paytype[2] == 'Internal_2' ? 'I' : '',
            PaytypeE: this.Paytype[3] == 'Extended_3' ? 'E' : '',

            DepartmentS: this.Department.indexOf('Service') >= 0 ? 'S' : '',
            DepartmentP: this.Department.indexOf('Parts') >= 0 ? 'P' : '',            // DepartmentP: '',
            DepartmentQ: this.Department.indexOf('Quicklube') >= 0 ? 'Q' : '',
            // DepartmentB: this.Department.indexOf('Details') >= 0 ? 'D' : '',
            DepartmentB: this.Department.indexOf('Body') >= 0 ? 'B' : '',
            DepartmentD: this.Department.indexOf('Details') >= 0 ? 'D' : '',
            PolicyAccount: this.labourType == 'Y' ? this.labourType : 'N',
            userName: Item.data1,
            Grosstype: this.Grosstype,
            layer: 1,
            zeroHours: this.zeroHours == true ? 'Y' : 'N',
            GrossTypeM: this.Misc == true ? 'M' : '',
            GrossTypeL: this.Lube == true ? 'L' : '',
            GrossTypeS: this.Sublet == true ? 'S' : '',
            GrossTypeP: this.Grosstype[1] == 'Parts_1' ? 'P' : '',
            LaborTypes: this.LaborTypeVal.toString()
          },
        ];
      }
    }
    // //console.log(Item)
    if (ref == '2') {
      if (Item.data2 != undefined) {
        const DetailsServicePerson = this.shared.ngbmodal.open(
          ServicegrossDetails,
          {
            // size:'xl',
            backdrop: 'static',
          }
        );
        DetailsServicePerson.componentInstance.Servicedetails = [
          {
            // "storeId": ParentItem.Store,
            //  "SrvcName": Item.ServiceAdvisor_Name,
            StartDate: this.FromDate,
            EndDate: this.ToDate,
            var1: this.selectedDataGrouping[0]?.columnname,
            var2: this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '',
            var3: '',
            var1Value: ParentItem.data1,
            var2Value: Item.data2,
            var3Value: '',
            PaytypeC: this.Paytype[0] == 'Customerpay_0' ? 'C' : '',
            PaytypeW: this.Paytype[1] == 'Warranty_1' ? 'W' : '',
            PaytypeI: this.Paytype[2] == 'Internal_2' ? 'I' : '',
            PaytypeE: this.Paytype[3] == 'Extended_3' ? 'E' : '',

            DepartmentS: this.Department.indexOf('Service') >= 0 ? 'S' : '',
            DepartmentP: this.Department.indexOf('Parts') >= 0 ? 'P' : '',            // DepartmentP: '',
            DepartmentQ: this.Department.indexOf('Quicklube') >= 0 ? 'Q' : '',
            // DepartmentB: this.Department.indexOf('Details') >= 0 ? 'D' : '',
            DepartmentB: this.Department.indexOf('Body') >= 0 ? 'B' : '',
            DepartmentD: this.Department.indexOf('Details') >= 0 ? 'D' : '',
            PolicyAccount: this.labourType == 'Y' ? this.labourType : 'N',
            userName: Item.data2,
            Grosstype: this.Grosstype,
            layer: 2,
            zeroHours: this.zeroHours == true ? 'Y' : 'N',
            GrossTypeM: this.Misc == true ? 'M' : '',
            GrossTypeL: this.Lube == true ? 'L' : '',
            GrossTypeS: this.Sublet == true ? 'S' : '',
            GrossTypeP: this.Grosstype[1] == 'Parts_1' ? 'P' : '',
            LaborTypes: this.LaborTypeVal.toString()

          },
        ];
      }
    }
    if (ref == '3') {
      if (Item.data3 != undefined) {
        const DetailsServicePerson = this.shared.ngbmodal.open(
          ServicegrossDetails,
          {
            // size:'xl',
            backdrop: 'static',
          }
        );
        DetailsServicePerson.componentInstance.Servicedetails = [
          {
            StartDate: this.FromDate,
            EndDate: this.ToDate,
            var1: this.selectedDataGrouping[0]?.columnname,
            var2: this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '',
            var3: (this.selectedDataGrouping.length == 3 ? this.selectedDataGrouping[2]?.columnname : ''),
            var1Value: ParentItem.data1,
            var2Value: Item.data2,
            var3Value: Item.data3,
            PaytypeC: this.Paytype[0] == 'Customerpay_0' ? 'C' : '',
            PaytypeW: this.Paytype[1] == 'Warranty_1' ? 'W' : '',
            PaytypeI: this.Paytype[2] == 'Internal_2' ? 'I' : '',
            PaytypeE: this.Paytype[3] == 'Extended_3' ? 'E' : '',

            DepartmentS: this.Department.indexOf('Service') >= 0 ? 'S' : '',
            DepartmentP: this.Department.indexOf('Parts') >= 0 ? 'P' : '',            // DepartmentP: '',
            DepartmentQ: this.Department.indexOf('Quicklube') >= 0 ? 'Q' : '',
            // DepartmentB: this.Department.indexOf('Details') >= 0 ? 'D' : '',
            DepartmentB: this.Department.indexOf('Body') >= 0 ? 'B' : '',
            DepartmentD: this.Department.indexOf('Details') >= 0 ? 'D' : '',
            PolicyAccount: this.labourType == 'Y' ? this.labourType : 'N',
            userName: Item.data3,
            Grosstype: this.Grosstype,
            layer: 3,
            zeroHours: this.zeroHours == true ? 'Y' : 'N',
            GrossTypeM: this.Misc == true ? 'M' : '',
            GrossTypeL: this.Lube == true ? 'L' : '',
            GrossTypeS: this.Sublet == true ? 'S' : '',
            GrossTypeP: this.Grosstype[1] == 'Parts_1' ? 'P' : '',
            LaborTypes: this.LaborTypeVal.toString()
          },
        ];
        DetailsServicePerson.result.then(
          (data) => { },
          (reason) => {
            // on dismiss
          }
        );
      }
    }
  }


  ngOnInit(): void {

    // this.GetSalesReconciliationData();

    // var curl = 'https://fbxtractapi.axelautomotive.com/api/fredbeans/GetServiceSummaryBetaDetailsV2';
    // this.shared.api.logSaving(curl, {}, '', 'Success', 'Service Gross');
  }


  ngAfterViewInit() {
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.comm.pageName == 'Service Gross') {
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


    this.excel = this.shared.api.getExportToExcelAllReports().subscribe((res) => {
      if (this.excel != undefined) {
        if (res.obj.title == 'Service Gross') {
          if (res.obj.state == true) {
            this.exportToExcel();
          }
        }
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res) => {
      if (this.print != undefined) {
        if (res.obj.title == 'Service Gross') {
          if (res.obj.statePrint == true) {
            //  this.GetPrintData();
          }
        }
      }
    });


    this.Pdf = this.shared.api.getExportToPDFAllReports().subscribe((res) => {
      if (this.Pdf != undefined) {
        if (res.obj.title == 'Service Gross') {
          if (res.obj.statePDF == true) {
            //  this.generatePDF();
          }
        }
      }
    });
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res) => {
      if (this.email != undefined) {
        if (res.obj.title == 'Service Gross') {
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

  PayTypeGrid(val: any) {
    console.log(this.getState, ' Get State');

    if (val == 'G') {
      this.GridView = 'Gross'
      this.headervalues();
      this.GetGrossData();
    }
    if (val == 'H') {
      this.GridView = 'Hours'
      this.headervalues();
      this.GetHoursData();
    }
    if (val == 'R') {
      this.GridView = 'RO'
      this.headervalues();
      this.GetROData();
    }
    // this.header = [{ type: 'Bar', storeIds: this.store, fromDate: this.FromDate, zeroHours: this.zeroHours, Sublet: this.Sublet, Misc: this.Misc, Lube: this.Lube, toDate: this.ToDate, ReportTotal: this.TotalReport, Paytype: this.Paytype, groups: this.groups, gridview: this.GridView, otherstoreids: this.otherstoreid, selectedotherstoreids: this.selectedotherstoreids }]

  }

  GetGrossData() {
    if (this.getState.indexOf('P') < 0) {
      this.responcestatus = '';
      this.shared.spinner.show();
      this.GetGrossIndiData();
      this.GetGrossTotalData();
    }
  }
  GetGrossIndiData() {
    this.GridView = 'Gross';
    this.headervalues()


    this.IndividualServicePaytype = [];
    this.shared.spinner.show();
    const obj = {
      startdate: this.FromDate,
      enddate: this.ToDate,
      StoreID: [...this.storeIds, ...this.otherStoreIds],
      AdvisorNumber: '',
      AdvisorName: '',
      ROSTATUS: '',
      PaytypeC: this.Paytype[0] == 'Customerpay_0' ? 'C' : '',
      PaytypeW: this.Paytype[1] == 'Warranty_1' ? 'W' : '',
      PaytypeI: this.Paytype[2] == 'Internal_2' ? 'I' : '',
      PaytypeE: this.Paytype[3] == 'Extended_3' ? 'E' : '',

      DepartmentS: this.Department.indexOf('Service') >= 0 ? 'S' : '',
      DepartmentP: this.Department.indexOf('Parts') >= 0 ? 'P' : '',
      // DepartmentP: '',
      DepartmentQ: this.Department.indexOf('Quicklube') >= 0 ? 'Q' : '',
      DepartmentB: this.Department.indexOf('Body') >= 0 ? 'B' : '',
      DepartmentD: this.Department.indexOf('Details') >= 0 ? 'D' : '',
      GrossTypeM: this.Misc == true ? 'M' : '',
      GrossTypeL: this.Lube == true ? 'L' : '',
      GrossTypeS: this.Sublet == true ? 'S' : '',
      GrossTypeP: '',
      PolicyAccount: this.labourType == 'Y' ? this.labourType : 'N',
      excludeZeroHours: this.zeroHours == true ? 'Y' : 'N',

      vehicle_Year: '',
      vehicle_Make: '',
      Vehicle_Model: '',
      Vehicle_Odometer: '',
      CName: '',
      CZip: '',
      CState: '',
      RO_CloseDate: '',
      var1: this.selectedDataGrouping.length >= 1 ? this.selectedDataGrouping[0]?.columnname : '',
      var2: this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '',
      var3: this.selectedDataGrouping.length == 3 ? this.selectedDataGrouping[2]?.columnname : '',
      type: 'D',
      LaborTypes: this.LaborTypeVal.toString()
    };

    this.shared.api
      .postmethod(this.comm.routeEndpoint + 'GetServiceSummaryBetaPaytypeV2Sublet', obj)
      .subscribe(
        (res) => {
          if (res.status == 200) {
            if (res.response != undefined) {
              if (res.response.length > 0) {
                this.IndividualServicePaytype = [];
                this.IndividualServicePaytype = res.response;
                this.responcestatus = this.responcestatus + 'I';
                let length = this.IndividualServicePaytype.length;
                let path2 = this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '';
                this.IndividualServicePaytype.some(function (x: any) {
                  x.Dealer = '+';
                  if (x.Data2 != undefined && x.Data2 != null) {
                    x.Data2 = JSON.parse(x.Data2);
                    x.Data2 = x.Data2.map((v: any) => ({
                      ...v,
                      SubData: [],
                      data2sign: '-',
                    }));
                  }
                  if (x.Data3 != undefined) {
                    x.Data3 = JSON.parse(x.Data3);
                    x.Data2.forEach((val: any) => {
                      x.Data3.forEach((ele: any) => {
                        if (val.data2 == ele.data2) {
                          val.SubData.push(ele);
                        }
                      });
                    });
                  }
                  if (path2 == '') {
                    x.Dealer = '+';
                  } else {
                    x.Dealer = '-';
                  }
                });
                this.combinePayType();
              } else {
                // this.toast.show('Empty Response','');
                this.shared.spinner.hide();
                this.NoData = true;
              }
            } else {
              // this.toast.show('Empty Response','');
              this.shared.spinner.hide();
              this.NoData = true;
            }
          } else {
            //this.toast.show('No Response', '');
            this.shared.spinner.hide();
            this.NoData = true;
          }
        },
        (error) => {
          this.toast.show('502 Bad Gate Way Error', 'danger', 'Error');

          this.shared.spinner.hide();
          this.NoData = true;
        }
      );
  }

  GetGrossTotalData() {
    this.TotalServicePaytype = [];
    const obj = {
      startdate: this.FromDate,
      enddate: this.ToDate,
      StoreID: [...this.storeIds, ...this.otherStoreIds],
      AdvisorNumber: '',
      AdvisorName: '',
      ROSTATUS: '',
      PaytypeC: this.Paytype[0] == 'Customerpay_0' ? 'C' : '',
      PaytypeW: this.Paytype[1] == 'Warranty_1' ? 'W' : '',
      PaytypeI: this.Paytype[2] == 'Internal_2' ? 'I' : '',
      PaytypeE: this.Paytype[3] == 'Extended_3' ? 'E' : '',

      DepartmentS: this.Department.indexOf('Service') >= 0 ? 'S' : '',
      DepartmentP: this.Department.indexOf('Parts') >= 0 ? 'P' : '',
      // DepartmentP: '',
      DepartmentQ: this.Department.indexOf('Quicklube') >= 0 ? 'Q' : '',
      DepartmentB: this.Department.indexOf('Body') >= 0 ? 'B' : '',
      DepartmentD: this.Department.indexOf('Details') >= 0 ? 'D' : '',
      GrossTypeM: this.Misc == true ? 'M' : '',
      GrossTypeL: this.Lube == true ? 'L' : '',
      GrossTypeS: this.Sublet == true ? 'S' : '',
      GrossTypeP: '',
      PolicyAccount: this.labourType == 'Y' ? this.labourType : 'N',
      excludeZeroHours: this.zeroHours == true ? 'Y' : 'N',
      vehicle_Year: '',
      vehicle_Make: '',
      Vehicle_Model: '',
      Vehicle_Odometer: '',
      CName: '',
      CZip: '',
      CState: '',
      RO_CloseDate: '',
      var1: this.selectedDataGrouping[0].columnname == 'Store_Name' ? '' : this.selectedDataGrouping[0].columnname,
      var2: this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '',
      var3: this.selectedDataGrouping.length == 3 ? this.selectedDataGrouping[2]?.columnname : '',
      type: 'T',
      LaborTypes: this.LaborTypeVal.toString()
    };
    this.shared.api
      .postmethod(this.comm.routeEndpoint + 'GetServiceSummaryBetaPaytypeV2Sublet', obj)
      .subscribe(
        (totalres) => {
          if (totalres.status == 200) {
            if (totalres.response != undefined) {
              if (totalres.response.length > 0) {
                this.responcestatus = this.responcestatus + 'T';
                this.TotalServicePaytype = totalres.response.map((v: any) => ({
                  ...v,
                  data1: 'REPORTS TOTAL',
                  Dealer: '+',
                  Data2: [],
                }));
                // //console.log(this.TotalServicePaytype);
                this.combinePayType();
              } else {
                // this.toast.show('Empty Response','');
                this.shared.spinner.hide();
                this.NoData = true;
              }
            } else {
              // this.toast.show('Empty Response','');
              this.shared.spinner.hide();
              this.NoData = true;
            }
          } else {
            //this.toast.show('No Response', '');
            this.shared.spinner.hide();
            this.NoData = true;
          }
        },
        (error) => {
          this.toast.show('502 Bad Gate Way Error', 'danger', 'Error');

          // //console.log(error);
          this.shared.spinner.hide();
          this.NoData = true;
        }
      );
  }

  combinePayType() {
    if (this.responcestatus == 'IT' || this.responcestatus == 'TI') {
      if (this.TotalReport == 'B') {
        this.IndividualServicePaytype.push(this.TotalServicePaytype[0]);
        this.ServicePaytype = this.IndividualServicePaytype;
      } else {
        this.IndividualServicePaytype.unshift(this.TotalServicePaytype[0]);
        this.ServicePaytype = this.IndividualServicePaytype;
      }
      //console.log(this.ServicePaytype);

    } else if (this.responcestatus == 'T') {
      this.ServicePaytype = this.TotalServicePaytype;
    } else if (this.responcestatus == 'I') {
      this.ServicePaytype = this.IndividualServicePaytype;
    } else {
      this.NoData = true;
    }
    this.shared.spinner.hide();
    this.getState = this.getState + 'P'

  }

  GetHoursData() {
    if (this.getState.indexOf('H') < 0) {
      this.responcestatus = '';
      this.shared.spinner.show();
      this.GetHoursIndiData();
      this.GetHoursTotalData();
    }
  }

  GetHoursIndiData() {
    this.GridView = 'Hours';
    this.headervalues();
    this.IndividualServiceHour = [];
    this.shared.spinner.show();
    const obj = {
      startdate: this.FromDate,
      enddate: this.ToDate,
      StoreID: [...this.storeIds, ...this.otherStoreIds],

      AdvisorNumber: '',
      AdvisorName: '',
      ROSTATUS: '',
      PaytypeC: this.Paytype[0] == 'Customerpay_0' ? 'C' : '',
      PaytypeW: this.Paytype[1] == 'Warranty_1' ? 'W' : '',
      PaytypeI: this.Paytype[2] == 'Internal_2' ? 'I' : '',
      PaytypeE: this.Paytype[3] == 'Extended_3' ? 'E' : '',

      DepartmentS: this.Department.indexOf('Service') >= 0 ? 'S' : '',
      DepartmentP: this.Department.indexOf('Parts') >= 0 ? 'P' : '',
      // DepartmentP: '',
      DepartmentQ: this.Department.indexOf('Quicklube') >= 0 ? 'Q' : '',
      DepartmentB: this.Department.indexOf('Body') >= 0 ? 'B' : '',
      DepartmentD: this.Department.indexOf('Details') >= 0 ? 'D' : '',
      GrossTypeM: this.Misc == true ? 'M' : '',
      GrossTypeL: this.Lube == true ? 'L' : '',
      GrossTypeS: this.Sublet == true ? 'S' : '',
      GrossTypeP: '',
      PolicyAccount: this.labourType == 'Y' ? this.labourType : 'N',
      excludeZeroHours: this.zeroHours == true ? 'Y' : 'N',

      vehicle_Year: '',
      vehicle_Make: '',
      Vehicle_Model: '',
      Vehicle_Odometer: '',
      CName: '',
      CZip: '',
      CState: '',
      RO_CloseDate: '',
      var1: this.selectedDataGrouping.length >= 1 ? this.selectedDataGrouping[0]?.columnname : '',
      var2: this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '',
      var3: this.selectedDataGrouping.length == 3 ? this.selectedDataGrouping[2]?.columnname : '',
      type: 'D',
      LaborTypes: this.LaborTypeVal.toString()
    };

    this.shared.api
      .postmethod(this.comm.routeEndpoint + 'GetServiceSummaryHoursPaytypeV2Sublet', obj)
      .subscribe(
        (res) => {
          if (res.status == 200) {
            if (res.response != undefined) {
              if (res.response.length > 0) {
                this.IndividualServiceHour = [];
                this.IndividualServiceHour = res.response;
                this.responcestatus = this.responcestatus + 'I';
                let length = this.IndividualServiceHour.length;
                let path2 = this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '';
                let path3 = this.selectedDataGrouping.length == 3 ? this.selectedDataGrouping[2]?.columnname : '';
                this.IndividualServiceHour.some(function (x: any) {
                  x.Dealer = '+';
                  if (x.Data2 != undefined && x.Data2 != null) {
                    x.Data2 = JSON.parse(x.Data2);
                    x.Data2 = x.Data2.map((v: any) => ({
                      ...v,
                      SubData: [],
                      data2sign: '-',
                    }));
                  }
                  if (x.Data3 != undefined) {
                    x.Data3 = JSON.parse(x.Data3);
                    x.Data2.forEach((val: any) => {
                      x.Data3.forEach((ele: any) => {
                        if (val.data2 == ele.data2) {
                          val.SubData.push(ele);
                        }
                      });
                    });
                  }
                  if (path2 == '') {
                    x.Dealer = '+';
                  } else {
                    x.Dealer = '-';
                  }
                });
                this.combineHoursData();
              } else {
                // this.toast.show('Empty Response','');
                this.shared.spinner.hide();
                this.NoData = true;
              }
            } else {
              // this.toast.show('Empty Response','');
              this.shared.spinner.hide();
              this.NoData = true;
            }
          } else {
            //this.toast.show('No Response', '');
            this.shared.spinner.hide();
            this.NoData = true;
          }
        },
        (error) => {
          this.toast.show('502 Bad Gate Way Error', 'danger', 'Error');

          this.shared.spinner.hide();
          this.NoData = true;
        }
      );
  }

  GetHoursTotalData() {
    this.TotalServiceHour = [];
    const obj = {
      startdate: this.FromDate,
      enddate: this.ToDate,
      StoreID: [...this.storeIds, ...this.otherStoreIds],
      AdvisorNumber: '',
      AdvisorName: '',
      ROSTATUS: '',
      PaytypeC: this.Paytype[0] == 'Customerpay_0' ? 'C' : '',
      PaytypeW: this.Paytype[1] == 'Warranty_1' ? 'W' : '',
      PaytypeI: this.Paytype[2] == 'Internal_2' ? 'I' : '',
      PaytypeE: this.Paytype[3] == 'Extended_3' ? 'E' : '',

      DepartmentS: this.Department.indexOf('Service') >= 0 ? 'S' : '',
      DepartmentP: this.Department.indexOf('Parts') >= 0 ? 'P' : '',
      // DepartmentP: '',
      DepartmentQ: this.Department.indexOf('Quicklube') >= 0 ? 'Q' : '',
      DepartmentB: this.Department.indexOf('Body') >= 0 ? 'B' : '',
      DepartmentD: this.Department.indexOf('Details') >= 0 ? 'D' : '',
      GrossTypeM: this.Misc == true ? 'M' : '',
      GrossTypeL: this.Lube == true ? 'L' : '',
      GrossTypeS: this.Sublet == true ? 'S' : '',
      GrossTypeP: '',
      PolicyAccount: this.labourType == 'Y' ? this.labourType : 'N',
      excludeZeroHours: this.zeroHours == true ? 'Y' : 'N',

      vehicle_Year: '',
      vehicle_Make: '',
      Vehicle_Model: '',
      Vehicle_Odometer: '',
      CName: '',
      CZip: '',
      CState: '',
      RO_CloseDate: '',
      var1: this.selectedDataGrouping[0].columnname == 'Store_Name' ? '' : this.selectedDataGrouping[0].columnname,
      var2: this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '',
      var3: this.selectedDataGrouping.length == 3 ? this.selectedDataGrouping[2]?.columnname : '',
      type: 'T',
      LaborTypes: this.LaborTypeVal.toString()
    };
    this.shared.api
      .postmethod(this.comm.routeEndpoint + 'GetServiceSummaryHoursPaytypeV2Sublet', obj)
      .subscribe(
        (totalres) => {
          if (totalres.status == 200) {
            if (totalres.response != undefined) {
              if (totalres.response.length > 0) {
                this.TotalServiceHour = totalres.response.map((v: any) => ({
                  ...v,
                  data1: 'REPORTS TOTAL',
                  Dealer: '+',
                  Data2: [],
                }));
                this.responcestatus = this.responcestatus + 'T';
                this.combineHoursData();
              } else {
                // this.toast.show('Empty Response','');
                this.shared.spinner.hide();
                this.NoData = true;
              }
            } else {
              // this.toast.show('Empty Response','');
              this.shared.spinner.hide();
              this.NoData = true;
            }
          } else {
            //this.toast.show('No Response', '');
            this.shared.spinner.hide();
            this.NoData = true;
          }
        },
        (error) => {
          this.toast.show('502 Bad Gate Way Error', 'danger', 'Error');

          //console.log(error);
          this.shared.spinner.hide();
          this.NoData = true;
        }
      );
  }
  combineHoursData() {
    if (this.responcestatus == 'IT' || this.responcestatus == 'TI') {
      if (this.TotalReport == 'B') {
        this.IndividualServiceHour.push(this.TotalServiceHour[0]);
        this.ServiceHour = this.IndividualServiceHour;
      } else {
        this.IndividualServiceHour.unshift(this.TotalServiceHour[0]);
        this.ServiceHour = this.IndividualServiceHour;
      }
    } else if (this.responcestatus == 'T') {
      this.ServiceHour = this.TotalServiceHour;
    } else if (this.responcestatus == 'I') {
      this.ServiceHour = this.IndividualServiceHour;
    } else {
      this.NoData = true;
    }
    this.shared.spinner.hide();
    this.getState = this.getState + 'H'

  }

  GetROData() {
    if (this.getState.indexOf('R') < 0) {
      this.responcestatus = '';
      this.shared.spinner.show();
      this.GetROIndiData();
      this.GetROTotalData();
    }
  }

  GetROIndiData() {
    this.GridView = 'RO';
    this.headervalues();
    this.IndividualServiceRO = [];
    this.shared.spinner.show();
    const obj = {
      startdate: this.FromDate,
      enddate: this.ToDate,
      StoreID: [...this.storeIds, ...this.otherStoreIds],
      AdvisorNumber: '',
      AdvisorName: '',
      ROSTATUS: '',
      PaytypeC: this.Paytype[0] == 'Customerpay_0' ? 'C' : '',
      PaytypeW: this.Paytype[1] == 'Warranty_1' ? 'W' : '',
      PaytypeI: this.Paytype[2] == 'Internal_2' ? 'I' : '',
      PaytypeE: this.Paytype[3] == 'Extended_3' ? 'E' : '',

      DepartmentS: this.Department.indexOf('Service') >= 0 ? 'S' : '',
      DepartmentP: this.Department.indexOf('Parts') >= 0 ? 'P' : '',
      // DepartmentP: '',
      DepartmentQ: this.Department.indexOf('Quicklube') >= 0 ? 'Q' : '',
      DepartmentB: this.Department.indexOf('Body') >= 0 ? 'B' : '',
      DepartmentD: this.Department.indexOf('Details') >= 0 ? 'D' : '',
      GrossTypeM: this.Misc == true ? 'M' : '',
      GrossTypeL: this.Lube == true ? 'L' : '',
      GrossTypeS: this.Sublet == true ? 'S' : '',
      GrossTypeP: '',
      PolicyAccount: this.labourType == 'Y' ? this.labourType : 'N',
      excludeZeroHours: this.zeroHours == true ? 'Y' : 'N',

      vehicle_Year: '',
      vehicle_Make: '',
      Vehicle_Model: '',
      Vehicle_Odometer: '',
      CName: '',
      CZip: '',
      CState: '',
      RO_CloseDate: '',
      var1: this.selectedDataGrouping.length >= 1 ? this.selectedDataGrouping[0]?.columnname : '',
      var2: this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '',
      var3: this.selectedDataGrouping.length == 3 ? this.selectedDataGrouping[2]?.columnname : '',
      type: 'D',
      LaborTypes: this.LaborTypeVal.toString()
    };

    this.shared.api
      .postmethod(this.comm.routeEndpoint + 'GetServiceSummaryROPaytypeV2Sublet', obj)
      .subscribe(
        (res) => {
          if (res.status == 200) {
            if (res.response != undefined) {
              if (res.response.length > 0) {
                this.IndividualServiceRO = [];
                this.IndividualServiceRO = res.response;
                let length = this.IndividualServiceRO.length;
                this.responcestatus = this.responcestatus + 'I';
                let path2 = this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '';
                let path3 = this.selectedDataGrouping.length == 3 ? this.selectedDataGrouping[2]?.columnname : '';
                this.IndividualServiceRO.some(function (x: any) {
                  x.Dealer = '+';
                  if (x.Data2 != undefined && x.Data2 != null) {
                    x.Data2 = JSON.parse(x.Data2);
                    x.Data2 = x.Data2.map((v: any) => ({
                      ...v,
                      SubData: [],
                      data2sign: '-',
                    }));
                  }
                  if (x.Data3 != undefined) {
                    x.Data3 = JSON.parse(x.Data3);
                    x.Data2.forEach((val: any) => {
                      x.Data3.forEach((ele: any) => {
                        if (val.data2 == ele.data2) {
                          val.SubData.push(ele);
                        }
                      });
                    });
                  }
                  if (path2 == '') {
                    x.Dealer = '+';
                  } else {
                    x.Dealer = '-';
                  }
                });
                this.combineROData();
              } else {
                // this.toast.show('Empty Response','');
                this.shared.spinner.hide();
                this.NoData = true;
              }
            } else {
              // this.toast.show('Empty Response','');
              this.shared.spinner.hide();
              this.NoData = true;
            }
          } else {
            //this.toast.show('No Response', '');
            this.shared.spinner.hide();
            this.NoData = true;
          }
        },
        (error) => {
          this.toast.show('502 Bad Gate Way Error', 'danger', 'Error');

          this.shared.spinner.hide();
          this.NoData = true;
        }
      );
  }
  GetROTotalData() {
    this.TotalServicePaytype = [];
    const obj = {
      startdate: this.FromDate,
      enddate: this.ToDate,
      StoreID: [...this.storeIds, ...this.otherStoreIds],
      AdvisorNumber: '',
      AdvisorName: '',
      ROSTATUS: '',
      PaytypeC: this.Paytype[0] == 'Customerpay_0' ? 'C' : '',
      PaytypeW: this.Paytype[1] == 'Warranty_1' ? 'W' : '',
      PaytypeI: this.Paytype[2] == 'Internal_2' ? 'I' : '',
      PaytypeE: this.Paytype[3] == 'Extended_3' ? 'E' : '',

      DepartmentS: this.Department.indexOf('Service') >= 0 ? 'S' : '',
      DepartmentP: this.Department.indexOf('Parts') >= 0 ? 'P' : '',
      // DepartmentP: '',
      DepartmentQ: this.Department.indexOf('Quicklube') >= 0 ? 'Q' : '',
      DepartmentB: this.Department.indexOf('Body') >= 0 ? 'B' : '',
      DepartmentD: this.Department.indexOf('Details') >= 0 ? 'D' : '',
      GrossTypeM: this.Misc == true ? 'M' : '',
      GrossTypeL: this.Lube == true ? 'L' : '',
      GrossTypeS: this.Sublet == true ? 'S' : '',
      GrossTypeP: '',
      PolicyAccount: this.labourType == 'Y' ? this.labourType : 'N',
      excludeZeroHours: this.zeroHours == true ? 'Y' : 'N',

      vehicle_Year: '',
      vehicle_Make: '',
      Vehicle_Model: '',
      Vehicle_Odometer: '',
      CName: '',
      CZip: '',
      CState: '',
      RO_CloseDate: '',
      var1: this.selectedDataGrouping[0].columnname == 'Store_Name' ? '' : this.selectedDataGrouping[0].columnname,
      var2: this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '',
      var3: this.selectedDataGrouping.length == 3 ? this.selectedDataGrouping[2]?.columnname : '',
      type: 'T',
      LaborTypes: this.LaborTypeVal.toString()
    };
    this.shared.api
      .postmethod(this.comm.routeEndpoint + 'GetServiceSummaryROPaytypeV2Sublet', obj)
      .subscribe(
        (totalres) => {
          if (totalres.status == 200) {
            if (totalres.response != undefined) {
              if (totalres.response.length > 0) {
                this.TotalServiceRO = totalres.response.map((v: any) => ({
                  ...v,
                  data1: 'REPORTS TOTAL',
                  Dealer: '+',
                  Data2: [],
                }));
                this.responcestatus = this.responcestatus + 'T';
                this.combineROData();
              } else {
                // this.toast.show('Empty Response','');
                this.shared.spinner.hide();
                this.NoData = true;
              }
            } else {
              // this.toast.show('Empty Response','');
              this.shared.spinner.hide();
              this.NoData = true;
            }
          } else {
            //this.toast.show('No Response', '');
            this.shared.spinner.hide();
            this.NoData = true;
          }
        },
        (error) => {
          this.toast.show('502 Bad Gate Way Error', 'danger', 'Error');

          //console.log(error);
          this.shared.spinner.hide();
          this.NoData = true;
        }
      );
  }

  combineROData() {
    if (this.responcestatus == 'IT' || this.responcestatus == 'TI') {
      if (this.TotalReport == 'B') {
        this.IndividualServiceRO.push(this.TotalServiceRO[0]);
        this.ServiceRO = this.IndividualServiceRO;
      } else {
        this.IndividualServiceRO.unshift(this.TotalServiceRO[0]);
        this.ServiceRO = this.IndividualServiceRO;
      }
    } else if (this.responcestatus == 'T') {
      this.ServiceRO = this.TotalServiceRO;
    } else if (this.responcestatus == 'I') {
      this.ServiceRO = this.IndividualServiceRO;
    } else {
      this.NoData = true;
    }
    console.log(this.ServiceRO)
    this.shared.spinner.hide();
    this.getState = this.getState + 'R'

  }
  back2grid() {
    this.GridView = 'Global';
    this.NoData = false
    this.headervalues();
    this.getServiceData();
  }

  StoresData(data: any) {
    this.storeIds = data.storeids;
    this.groupId = data.groupId;
    this.storename = data.storename;
    this.groupName = data.groupName;
    this.storecount = data.storecount;
    this.otherStoreIds = data.otherStoreIds;
    this.storedisplayname = data.storedisplayname;


    this.getlabourTypesData('FR', this.labourType)

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
      otherStoreIds: this.otherStoreIds, 'DefaultLoad': this.DefaultLoad
    };
  }
  updatedDates(data: any) {
    // console.log(data);
    this.FromDate = data.FromDate;
    this.ToDate = data.ToDate;
    this.DateType = data.DateType;
    this.displaytime = data.DisplayTime
  }

  setDates(type: any) {
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
  paytypeDisplay: any = ['Customerpay_0', 'Warranty_1', 'Internal_2', 'Extended_3'];
  SelectedData(val: any) {
    const index = this.selectedDataGrouping.findIndex((i: any) => i == val);
    if (index >= 0) {
      this.selectedDataGrouping.splice(index, 1);
    } else {
      if (this.selectedDataGrouping.length >= 3) {
        this.toast.show('Select up to 3 Filters only to Group your data', 'warning', 'Warning');
      } else {
        this.selectedDataGrouping.push(val);
      }
    }
  }
  multipleorsingle(block: any, e: any) {

    if (block == 'Dept') {
      const index = this.Department.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.Department.splice(index, 1);
        if (this.Department.length == 0) {
          this.toast.show('Please Select Atleast One Department', 'warning', 'Warning');
        }
      } else {
        this.Department.push(e);
      }

    }


    if (block == 'TB') {
      this.TotalReport = e;
    }
    if (block == 'LT') {
      this.labourType = e
      this.getlabourTypesData('FR', e)
    }

    if (block == 'PT') {
      let spliting = e.split('_');
      const index = this.Paytype.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.Paytype[spliting[1]] = '';
        let arr = this.Paytype.filter((e: any) => e != '');
        if (arr.length == 0) {
          this.toast.show('Please Select Atleast One Pay Type', 'warning', 'Warning');
        }
      } else {
        this.Paytype[spliting[1]] = e;
      }
      this.paytypeDisplay = this.Paytype.filter((val: any) => val != '');

    }
    if (block == 'GT') {
      let spliting = e.split('_');
      const index = this.Grosstype.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.Grosstype[spliting[1]] = '';
        let arr = this.Grosstype.filter((e: any) => e != '');
        if (arr.length == 0) {
          this.toast.show('Please Select Atleast One Gross Type', 'warning', 'Warning');
        }
      } else {
        this.Grosstype[spliting[1]] = e;
      }

    }
  }

  // alllabortypes() {
  //   if (this.LaborTypeVal.length == this.labortypes.length) {
  //     this.LaborTypeVal = []
  //   } else {
  //     this.LaborTypeVal = this.labortypes.map(function (a: any) {
  //       return a.ASD_labortype;
  //     });
  //   }
  // }

  alllabortypes(type: any) {
    if (type == 'Y') {
      this.LaborTypeVal = this.labortypes.map(function (a: any) {
        return a.ASD_labortype;
      });
    }
    else {
      this.LaborTypeVal = [];
    }
  }

  individualLabortypes(e: any) {
    const index = this.LaborTypeVal.findIndex(
      (i: any) => i == e.ASD_labortype
    );
    if (index >= 0) {
      this.LaborTypeVal.splice(index, 1);
    } else {
      this.LaborTypeVal.push(e.ASD_labortype);
    }
  }
  includeValues: any = []
  includes(e: any) {
    let truevalues: any = [this.zeroHours, this.Sublet, this.Misc, this.Lube]
    this.includeValues = truevalues.filter((val: any) => val == true)

  }
  spinnerLoaderlabor: boolean = false;
  activePopover: number = -1;

  togglePopover(popoverIndex: number) {
    this.activePopover = this.activePopover === popoverIndex ? -1 : popoverIndex;
  }
  // viewreport() {
  //   this.activePopover = -1
  //   if (this.selectedDataGrouping.length == 0) {
  //     this.toast.show('Please Select Atleast One Value from Grouping', 'warning', 'Warning');
  //   } else {
  //     if (this.storeIds.length == 0 && this.otherStoreIds.length == 0) {
  //       this.toast.show('Please Select Atleast One Store', 'warning', 'Warning');
  //     } else {
  //       let gt = this.Grosstype.filter((e: any) => e != '');
  //       let pt = this.Paytype.filter((e: any) => e != '');
  //       if (pt.length == 0 || gt.length == 0 || this.Department.length == 0) {

  //         if (this.Department.length == 0) {
  //           this.toast.show('Please Select Atleast One Department Type', 'warning', 'Warning');
  //         }
  //         if (pt.length == 0) {
  //           this.toast.show('Please Select Atleast One Pay Type', 'warning', 'Warning');
  //         }
  //         if (gt.length == 0) {
  //           this.toast.show('Please Select Atleast One Gross Type', 'warning', 'Warning');
  //         }
  //       } else {
  //         this.getState = ''
  //         this.headervalues();
  //         if (this.GridView == 'Global') {
  //           this.getServiceData();
  //         } else if (this.GridView == 'Gross') {
  //           this.GetGrossData();
  //         } else if (this.GridView == 'Hours') {
  //           this.GetHoursData();
  //         } else if (this.GridView == 'RO') {
  //           this.GetROData();
  //         }
  //       }
  //     }
  //   }
  // }
  viewreport() {
    this.activePopover = -1;
    if (!this.selectedDataGrouping || this.selectedDataGrouping.length === 0) {
      return this.toast.show('Please Select Atleast One Grouping', 'warning', 'Warning');
    }
    if ((!this.storeIds || this.storeIds.length === 0) &&
      (!this.otherStoreIds || this.otherStoreIds.length === 0)) {
      return this.toast.show('Please Select Atleast One Store', 'warning', 'Warning');
    }
    const gt = this.Grosstype?.filter((e: any) => e);
    const pt = this.Paytype?.filter((e: any) => e);
    if (!this.Department || this.Department.length === 0) {
      return this.toast.show('Please Select Atleast One Department Type', 'warning', 'Warning');
    }
    if (!pt || pt.length === 0) {
      return this.toast.show('Please Select Atleast One Pay Type', 'warning', 'Warning');
    }
    if (!gt || gt.length === 0) {
      return this.toast.show('Please Select Atleast One Gross Type', 'warning', 'Warning');
    }
    this.getState = '';
    this.headervalues();

    switch (this.GridView) {
      case 'Global':
        this.getServiceData();
        break;
      case 'Gross':
        this.GetGrossData();
        break;
      case 'Hours':
        this.GetHoursData();
        break;
      case 'RO':
        this.GetROData();
        break;
    }
  }

  ExcelStoreNames: any = [];
  exportToExcel() {

    let storeNames: any = [];
    let store = this.storeIds
    // const obj = {
    //   id: this.groups,
    //   userid: JSON.parse(localStorage.getItem('UserDetails')!).userid,
    // };
    // this.shared.api
    //   .postmethodOne('cavender/GetStoresbyGroupuserid', obj)
    //   .subscribe((res: any) => {
    storeNames = this.comm.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores
      .filter((item: any) =>
        store.some((cat: any) => cat === item.ID.toString())
      );
    if (store.length == this.comm.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores
      .length) {
      this.ExcelStoreNames = 'All Stores'
    } else {
      this.ExcelStoreNames = storeNames.map(function (a: any) {
        return a.storename;
      });
    }
    const ServiceData = this.ServiceData.map((_arrayElement: any) =>
      Object.assign({}, _arrayElement)
    );

    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('Service Gross');
    worksheet.views = [
      {
        showGridLines: false,
      },
    ];
    worksheet.addRow('');
    const titleRow = worksheet.addRow(['Service Gross']);
    titleRow.eachCell((cell: any, number: any) => {
      cell.alignment = { indent: 1, vertical: 'top', horizontal: 'left' };
    });
    titleRow.font = { name: 'Arial', family: 4, size: 12, bold: true };
    titleRow.worksheet.mergeCells('A2', 'D2');
    worksheet.addRow('');
    const PresentMonth = this.shared.datePipe.transform(this.FromDate, 'MMMM');
    const PresentYear = this.shared.datePipe.transform(this.FromDate, 'yyyy');
    const FromDate = this.shared.datePipe.transform(this.FromDate, 'dd');
    const ToDate = this.shared.datePipe.transform(this.ToDate, 'dd');
    const DateToday = this.shared.datePipe.transform(
      new Date(),
      'MM/dd/yyyy h:mm:ss a'
    );
    const DATE_EXTENSION = this.shared.datePipe.transform(
      new Date(),
      'MMddyyyy'
    );
    worksheet.addRow([DateToday]).font = { name: 'Arial', family: 4, size: 9 };

    const ReportFilter = worksheet.addRow(['Report Controls :']);
    ReportFilter.font = { name: 'Arial', family: 4, size: 10, bold: true };

    const Groupings = worksheet.addRow(['Groupings :']);
    Groupings.getCell(1).font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
    };
    const groupings = worksheet.getCell('B6');
    groupings.value =
      this.selectedDataGrouping[0]?.ARG_LABEL + ',' + this.selectedDataGrouping[1]?.ARG_LABEL + ',' + this.selectedDataGrouping[2]?.ARG_LABEL;
    groupings.font = { name: 'Arial', family: 4, size: 9 };

    const Timeframe = worksheet.addRow(['Timeframe :']);
    Timeframe.getCell(1).font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
    };
    const timeframe = worksheet.getCell('B7');
    timeframe.value = this.FromDate + ' to ' + this.ToDate;
    timeframe.font = { name: 'Arial', family: 4, size: 9 };

    const Stores = worksheet.addRow(['Stores :']);
    Stores.getCell(1).font = { name: 'Arial', family: 4, size: 9, bold: true };
    const Groups = worksheet.getCell('B8');
    Groups.value = 'Groups :';
    const groups = worksheet.getCell('D8');
    groups.value = this.comm.groupsandstores.filter((val: any) => val.sg_id == this.groupId.toString())[0].sg_name;

    groups.font = { name: 'Arial', family: 4, size: 9 };
    const Brands = worksheet.getCell('B9');
    Brands.value = 'Brands :';
    const brands = worksheet.getCell('D9');
    brands.value = '-';
    brands.font = { name: 'Arial', family: 4, size: 9 };
    const Stores1 = worksheet.getCell('B10');
    Stores1.value = 'Stores :';
    const stores1 = worksheet.getCell('D10');
    stores1.value = this.ExcelStoreNames == 0
      ? 'All Stores'
      : this.ExcelStoreNames == null
        ? '-'
        : this.ExcelStoreNames.toString().replaceAll(',', ', ');
    stores1.font = { name: 'Arial', family: 4, size: 9 };

    const Filters = worksheet.addRow(['Filters :']);
    Filters.getCell(1).font = { name: 'Arial', family: 4, size: 9, bold: true };
    const ROType = worksheet.getCell('B11');
    ROType.value = 'RO Type :';
    const rotype = worksheet.getCell('D11');
    rotype.value = '';
    rotype.font = { name: 'Arial', family: 4, size: 9 };
    const department = worksheet.getCell('B12');
    department.value = 'Department :';
    const Department = worksheet.getCell('D12');
    Department.value = '';
    Department.font = { name: 'Arial', family: 4, size: 9 };
    const PayType = worksheet.getCell('B13');
    PayType.value = 'Pay Type :';
    const paytype = worksheet.getCell('D13');
    paytype.value =
      this.Paytype[0] + ',' + this.Paytype[1] + ',' + this.Paytype[2];
    paytype.font = { name: 'Arial', family: 4, size: 9 };
    const SelectTarget = worksheet.getCell('B14');
    SelectTarget.value = 'Select Target :';
    const selecttarget = worksheet.getCell('D14');
    selecttarget.value = '';
    selecttarget.font = { name: 'Arial', family: 4, size: 9 };
    const GrossType = worksheet.getCell('B15');
    GrossType.value = 'Gross Type :';
    const grosstype = worksheet.getCell('D15');
    grosstype.value =
      this.Grosstype[0] +
      ',' +
      this.Grosstype[1] +
      ',' +
      this.Grosstype[2] +
      ',' +
      this.Grosstype[3];
    grosstype.font = { name: 'Arial', family: 4, size: 9 };
    const Source = worksheet.getCell('B16');
    Source.value = 'Source :';
    const source = worksheet.getCell('D16');
    source.value = '';
    source.font = { name: 'Arial', family: 4, size: 9 };
    const ReportTotals = worksheet.getCell('B17');
    ReportTotals.value = 'Report Totals :';
    const reporttotals = worksheet.getCell('D17');
    reporttotals.value = this.TotalReport == 'T' ? 'Top' : 'Bottom';
    reporttotals.font = { name: 'Arial', family: 4, size: 9 };
    worksheet.addRow('');

    let dateYear = worksheet.getCell('A19');
    dateYear.value = PresentMonth;
    dateYear.alignment = { vertical: 'middle', horizontal: 'center' };
    dateYear.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
      color: { argb: 'FFFFFF' },
    };
    dateYear.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '2a91f0' },
      bgColor: { argb: 'FF0000FF' },
    };
    dateYear.border = { right: { style: 'dotted' } };

    worksheet.mergeCells('B19', 'G19');
    let units = worksheet.getCell('B19');
    units.value = 'Gross';
    units.alignment = { vertical: 'middle', horizontal: 'center' };
    units.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
      color: { argb: 'FFFFFF' },
    };
    units.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '2a91f0' },
      bgColor: { argb: 'FF0000FF' },
    };
    units.border = { right: { style: 'dotted' } };

    worksheet.mergeCells('H19', 'K19');
    let frontgross = worksheet.getCell('H19');
    frontgross.value = 'Hours';
    frontgross.alignment = { vertical: 'middle', horizontal: 'center' };
    frontgross.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
      color: { argb: 'FFFFFF' },
    };
    frontgross.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '2a91f0' },
      bgColor: { argb: 'FF0000FF' },
    };
    frontgross.border = { right: { style: 'dotted' } };

    worksheet.mergeCells('L19', 'O19');
    let backgross = worksheet.getCell('L19');
    backgross.value = 'Repair Orders';
    backgross.alignment = { vertical: 'middle', horizontal: 'center' };
    backgross.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
      color: { argb: 'FFFFFF' },
    };
    backgross.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '2a91f0' },
      bgColor: { argb: 'FF0000FF' },
    };
    backgross.border = { right: { style: 'dotted' } };

    worksheet.mergeCells('P19', 'V19');
    let totalgross = worksheet.getCell('P19');
    totalgross.value = 'Performance';
    totalgross.alignment = { vertical: 'middle', horizontal: 'center' };
    totalgross.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
      color: { argb: 'FFFFFF' },
    };
    totalgross.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '2a91f0' },
      bgColor: { argb: 'FF0000FF' },
    };
    totalgross.border = { right: { style: 'thin' } };

    let Headings = [
      FromDate + ' - ' + ToDate + ', ' + PresentYear,
      'Service Sale',
      'Total',
      'Pace',
      'Target',
      'Diff',
      'Discounts',
      'Total',
      'Pace',
      'Target',
      'Diff',
      'Total',
      'Pace',
      'Target',
      'Diff',
      'Hours/RO	',
      'Sales/RO',
      'Parts/RO	',
      'Avg RO',
      'ELR',
      'GP%',
      'MPI%',
    ];
    const headerRow = worksheet.addRow(Headings);
    headerRow.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: false,
      color: { argb: 'Black' },
    };
    headerRow.alignment = { indent: 1, vertical: 'top', horizontal: 'center' };
    headerRow.eachCell((cell: any, number: any) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'D9D9D9' },
        bgColor: { argb: 'FF0000FF' },
      };
      cell.border = { right: { style: 'thin' } };
      cell.alignment = { vertical: 'top', horizontal: 'center' };
    });

    for (const d of ServiceData) {
      const Data1 = worksheet.addRow([
        d.data1 == '' ? '-' : d.data1 == null ? '-' : d.data1,
        d.TotalSale == '' ? '-' : d.TotalSale == null ? '-' : d.TotalSale,
        d.Total_Gross == '' ? '-' : d.Total_Gross == null ? '-' : d.Total_Gross,
        d.TotalGross_Pace == '' || this.shared.datePipe.transform(this.FromDate, 'MMMM') != this.shared.datePipe.transform(this.ToDate, 'MMMM')
          ? '-'
          : d.TotalGross_Pace == null
            ? '-'
            : d.TotalGross_Pace,
        d.Gross_Target == '' || this.shared.datePipe.transform(this.FromDate, 'MMMM') != this.shared.datePipe.transform(this.ToDate, 'MMMM')
          ? '-'
          : d.Gross_Target == null
            ? '-'
            : d.Gross_Target,
        d.Diff == '' || this.shared.datePipe.transform(this.FromDate, 'MMMM') != this.shared.datePipe.transform(this.ToDate, 'MMMM') ? '-' : d.Diff == null ? '-' : d.Diff,
        d.Discount == '' ? '-' : d.Discount == null ? '-' : d.Discount,
        d.Total_Hours == '' ? '-' : d.Total_Hours == null ? '-' : d.Total_Hours,
        d.TotalHours_PACE == '' || this.shared.datePipe.transform(this.FromDate, 'MMMM') != this.shared.datePipe.transform(this.ToDate, 'MMMM')
          ? '-'
          : d.TotalHours_PACE == null
            ? '-'
            : d.TotalHours_PACE,
        d.ROHours_Target == '' || this.shared.datePipe.transform(this.FromDate, 'MMMM') != this.shared.datePipe.transform(this.ToDate, 'MMMM')
          ? '-'
          : d.ROHours_Target == null
            ? '-'
            : d.ROHours_Target,
        d.HoursDiff == '' || this.shared.datePipe.transform(this.FromDate, 'MMMM') != this.shared.datePipe.transform(this.ToDate, 'MMMM') ? '-' : d.HoursDiff == null ? '-' : d.HoursDiff,
        d.Repair_Orders == ''
          ? '-'
          : d.Repair_Orders == null
            ? '-'
            : d.Repair_Orders,
        d.Total_ROPACE == '' || this.shared.datePipe.transform(this.FromDate, 'MMMM') != this.shared.datePipe.transform(this.ToDate, 'MMMM')
          ? '-'
          : d.Total_ROPACE == null
            ? '-'
            : d.Total_ROPACE,
        d.RO_Target == '' || this.shared.datePipe.transform(this.FromDate, 'MMMM') != this.shared.datePipe.transform(this.ToDate, 'MMMM') ? '-' : d.RO_Target == null ? '-' : d.RO_Target,
        d.ROCountDiff == '' || this.shared.datePipe.transform(this.FromDate, 'MMMM') != this.shared.datePipe.transform(this.ToDate, 'MMMM') ? '-' : d.ROCountDiff == null ? '-' : d.ROCountDiff,
        d.Hours_per_RO == ''
          ? '-'
          : d.Hours_per_RO == null
            ? '-'
            : d.Hours_per_RO,
        d.Sales_Per_RO == ''
          ? '-'
          : d.Sales_Per_RO == null
            ? '-'
            : d.Sales_Per_RO,
        d.Parts_Per_RO == ''
          ? '-'
          : d.Parts_Per_RO == null
            ? '-'
            : d.Parts_Per_RO,
        d.Average_RO == '' ? '-' : d.Average_RO == null ? '-' : d.Average_RO,
        d.ELR == '' ? '-' : d.ELR == null ? '-' : d.ELR,
        d.Retention == ''
          ? '-'
          : d.Retention == null
            ? '-'
            : d.Retention + ' %',
        d.MPI == '' ? '-' : d.MPI == null ? '-' : d.MPI + ' %',
      ]);
      // Data1.outlineLevel = 1; // Grouping level 1
      Data1.font = { name: 'Arial', family: 4, size: 9 };
      Data1.getCell(1).alignment = {
        indent: 1,
        vertical: 'top',
        horizontal: 'left',
      };
      Data1.eachCell((cell: any, number: any) => {
        cell.border = { right: { style: 'thin' } };
        if (
          (number > 1 && number < 8) ||
          number == 17 ||
          number == 19
        ) {
          cell.numFmt = '$#,##0';
        }
        if (number > 8 && number < 16) {
          cell.numFmt = '#,##0';
        }
        if (number == 16 || number == 21 || number == 22 ||
          number == 20) {
          cell.numFmt = '#,##0.00';
        }
        if (number == 18 || number == 8) {
          cell.numFmt = '#,##0.0';
        }
        if (number != 1) {
          cell.alignment = { vertical: 'top', horizontal: 'right', indent: 1 };
        }
      });
      if (Data1.number % 2) {
        Data1.eachCell((cell, number) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'e5e5e5' },
            bgColor: { argb: 'FF0000FF' },
          };
        });
      }
      if (d.Data2 != undefined) {
        for (const d1 of d.Data2) {
          const Data2 = worksheet.addRow([
            d1.data2 == '' ? '-' : d1.data2 == null ? '-' : d1.data2,
            d1.TotalSale == ''
              ? '-'
              : d1.TotalSale == null
                ? '-'
                : d1.TotalSale,
            d1.Total_Gross == ''
              ? '-'
              : d1.Total_Gross == null
                ? '-'
                : d1.Total_Gross,
            d1.TotalGross_Pace == '' || this.shared.datePipe.transform(this.FromDate, 'MMMM') != this.shared.datePipe.transform(this.ToDate, 'MMMM')
              ? '-'
              : d1.TotalGross_Pace == null
                ? '-'
                : d1.TotalGross_Pace,
            '-',
            '-',
            d1.Discount == '' ? '-' : d1.Discount == null ? '-' : d1.Discount,

            d1.Total_Hours == ''
              ? '-'
              : d1.Total_Hours == null
                ? '-'
                : d1.Total_Hours,
            d1.TotalHours_PACE == '' || this.shared.datePipe.transform(this.FromDate, 'MMMM') != this.shared.datePipe.transform(this.ToDate, 'MMMM')
              ? '-'
              : d1.TotalHours_PACE == null
                ? '-'
                : d1.TotalHours_PACE,
            '-',
            '-',

            d1.Repair_Orders == ''
              ? '-'
              : d1.Repair_Orders == null
                ? '-'
                : d1.Repair_Orders,
            d1.Total_ROPACE == ''
              ? '-'
              : d1.Total_ROPACE == null
                ? '-'
                : d1.Total_ROPACE,
            '-',
            '-',

            d1.Hours_per_RO == ''
              ? '-'
              : d1.Hours_per_RO == null
                ? '-'
                : d1.Hours_per_RO,
            d1.Sales_Per_RO == ''
              ? '-'
              : d1.Sales_Per_RO == null
                ? '-'
                : d1.Sales_Per_RO,
            d1.Parts_Per_RO == ''
              ? '-'
              : d1.Parts_Per_RO == null
                ? '-'
                : d1.Parts_Per_RO,
            d1.Average_RO == ''
              ? '-'
              : d1.Average_RO == null
                ? '-'
                : d1.Average_RO,
            d1.ELR == '' ? '-' : d1.ELR == null ? '-' : d1.ELR,
            d1.Retention == ''
              ? '-'
              : d1.Retention == null
                ? '-'
                : d1.Retention + ' %',
            d1.MPI == '' ? '-' : d1.MPI == null ? '-' : d1.MPI + ' %',
          ]);
          Data2.outlineLevel = 1; // Grouping level 2
          Data2.font = { name: 'Arial', family: 4, size: 8 };
          Data2.getCell(1).alignment = {
            indent: 2,
            vertical: 'top',
            horizontal: 'left',
          };
          Data2.eachCell((cell: any, number: any) => {
            cell.border = { right: { style: 'thin' } };
            if (
              (number > 1 && number < 8) ||
              number == 17 ||
              number == 19

            ) {
              cell.numFmt = '$#,##0';
            }
            if (number > 8 && number < 16) {
              cell.numFmt = '#,##0';
            }
            if (number == 16 || number == 21 || number == 22 || number == 20) {
              cell.numFmt = '#,##0.00';
            }
            if (number == 18 || number == 8) {
              cell.numFmt = '#,##0.0';
            }
            if (number != 1) {
              cell.alignment = {
                vertical: 'top',
                horizontal: 'right',
                indent: 1,
              };
            }
          });
          if (Data2.number % 2) {
            Data2.eachCell((cell, number) => {
              cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'e5e5e5' },
                bgColor: { argb: 'FF0000FF' },
              };
            });
          }
          if (d1.SubData != undefined) {
            for (const d2 of d1.SubData) {
              const Data3 = worksheet.addRow([
                d2.data3 == '' ? '-' : d2.data3 == null ? '-' : d2.data3,
                d2.TotalSale == ''
                  ? '-'
                  : d2.TotalSale == null
                    ? '-'
                    : d2.TotalSale,
                d2.Total_Gross == ''
                  ? '-'
                  : d2.Total_Gross == null
                    ? '-'
                    : d2.Total_Gross,
                d2.TotalGross_Pace == '' || this.shared.datePipe.transform(this.FromDate, 'MMMM') != this.shared.datePipe.transform(this.ToDate, 'MMMM')
                  ? '-'
                  : d2.TotalGross_Pace == null
                    ? '-'
                    : d2.TotalGross_Pace,
                '-',
                '-',
                d2.Discount == '' ? '-' : d2.Discount == null ? '-' : d2.Discount,

                d2.Total_Hours == ''
                  ? '-'
                  : d2.Total_Hours == null
                    ? '-'
                    : d2.Total_Hours,
                d2.TotalHours_PACE == '' || this.shared.datePipe.transform(this.FromDate, 'MMMM') != this.shared.datePipe.transform(this.ToDate, 'MMMM')
                  ? '-'
                  : d2.TotalHours_PACE == null
                    ? '-'
                    : d2.TotalHours_PACE,
                '-',
                '-',

                d2.Repair_Orders == ''
                  ? '-'
                  : d2.Repair_Orders == null
                    ? '-'
                    : d2.Repair_Orders,
                d2.RO_PACE == '' || this.shared.datePipe.transform(this.FromDate, 'MMMM') != this.shared.datePipe.transform(this.ToDate, 'MMMM')
                  ? '-'
                  : d2.RO_PACE == null
                    ? '-'
                    : d2.RO_PACE,
                '-',
                '-',

                d2.Hours_per_RO == ''
                  ? '-'
                  : d2.Hours_per_RO == null
                    ? '-'
                    : d2.Hours_per_RO,
                d2.Sales_Per_RO == ''
                  ? '-'
                  : d2.Sales_Per_RO == null
                    ? '-'
                    : d2.Sales_Per_RO,
                d2.Parts_Per_RO == ''
                  ? '-'
                  : d2.Parts_Per_RO == null
                    ? '-'
                    : d2.Parts_Per_RO,

                d2.Average_RO == ''
                  ? '-'
                  : d2.Average_RO == null
                    ? '-'
                    : d2.Average_RO,
                d2.ELR == '' ? '-' : d2.ELR == null ? '-' : d2.ELR,
                d2.Retention == ''
                  ? '-'
                  : d2.Retention == null
                    ? '-'
                    : d2.Retention + ' %',
                d2.MPI == '' ? '-' : d2.MPI == null ? '-' : d2.MPI + ' %',
              ]);
              Data3.outlineLevel = 2; // Grouping level 2
              Data3.font = { name: 'Arial', family: 4, size: 8 };
              Data3.alignment = {
                vertical: 'middle',
                horizontal: 'center',
              };
              Data3.getCell(1).alignment = {
                indent: 3,
                vertical: 'middle',
                horizontal: 'left',
              };
              Data3.eachCell((cell: any, number: any) => {
                cell.border = { right: { style: 'dotted' } };
                cell.numFmt = '$#,##0';
                if (
                  (number > 1 && number < 8) ||
                  number == 17 ||
                  number == 19
                ) {
                  cell.numFmt = '$#,##0';
                }
                if (number > 8 && number < 16) {
                  cell.numFmt = '#,##0';
                }
                if (number == 16 || number == 21 || number == 22 ||
                  number == 20) {
                  cell.numFmt = '#,##0.00';
                }
                if (number == 18 || number == 8) {
                  cell.numFmt = '#,##0.0';
                }
                if (number != 1) {
                  cell.alignment = {
                    vertical: 'top',
                    horizontal: 'right',
                    indent: 1,
                  };
                }
              });
              if (Data3.number % 2) {
                Data3.eachCell((cell, number) => {
                  cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'e5e5e5' },
                    bgColor: { argb: 'FF0000FF' },
                  };
                });
              }

            }
          }
        }
      }
    }
    worksheet.eachRow((row, rowIndex) => {
      row.eachCell((cell, colIndex) => {
        if (rowIndex > 1 && rowIndex < 19) { // Skip the header row
          // Apply conditional alignment based on your conditions
          if (colIndex === 1) {
            // Apply right alignment to the second column
            cell.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
          }
        }
      });
    });
    worksheet.getColumn(1).width = 30;
    worksheet.getColumn(2).width = 15;
    worksheet.getColumn(3).width = 15;
    worksheet.getColumn(4).width = 15;
    worksheet.getColumn(5).width = 15;
    worksheet.getColumn(6).width = 15;
    worksheet.getColumn(7).width = 15;
    worksheet.getColumn(8).width = 15;
    worksheet.getColumn(9).width = 15;
    worksheet.getColumn(10).width = 15;
    worksheet.getColumn(11).width = 15;
    worksheet.getColumn(12).width = 15;
    worksheet.getColumn(13).width = 15;
    worksheet.getColumn(14).width = 15;
    worksheet.getColumn(15).width = 15;
    worksheet.getColumn(16).width = 15;
    worksheet.getColumn(17).width = 15;
    worksheet.getColumn(18).width = 15;
    worksheet.getColumn(19).width = 15;
    worksheet.getColumn(20).width = 15;
    worksheet.getColumn(21).width = 15;
    worksheet.addRow([]);

    this.shared.exportToExcel(workbook, 'Service Gross V2_' + DATE_EXTENSION);

  }


}