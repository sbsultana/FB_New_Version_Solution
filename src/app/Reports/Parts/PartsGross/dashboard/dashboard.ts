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
import { PartsGrossDetails } from '../parts-gross-details/parts-gross-details';

@Component({
  selector: 'app-dashboard',
  imports: [SharedModule, BsDatepickerModule, DateRangePicker, Stores],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
})
export class Dashboard {

  PartsData: any = [];
  IndividualPartsGross: any = [];
  TotalPartsGross: any = [];
  TotalReport: any = 'T';
  NoData: boolean = false;
  Department: any = ['Parts', 'Service'];
  Paytype: any = ['R', 'W'];
  partsSource: any = 'All';

  responcestatus: any = '';
  groups: any = 1;
  servicetype: any = ['C', 'T', 'I', 'E']
  GridView = 'Global';
  otherstoreid: any = '';
  selectedotherstoreids: any = ''

  selectedDataGrouping: any = [];
  dataGrouping: any = [    
      { "ARG_ID": 48, "ARG_LABEL": "Store", "ARG_SEQ": 0, "id": 48, "columnname": "DealerName", "Active": "Y" },
      { "ARG_ID": 41, "ARG_LABEL": "Counterperson", "ARG_SEQ": 1, "id": 41, "columnname": "AP_CounterPerson", "Active": "Y" },
      { "ARG_ID": 42, "ARG_LABEL": "Sale Type", "ARG_SEQ": 2, "id": 42, "columnname": "AP_LP_Type", "Active": "Y" },
      { "ARG_ID": 43, "ARG_LABEL": "Customer Name", "ARG_SEQ": 3, "id": 43, "columnname": "Customername", "Active": "Y" },
      { "ARG_ID": 44, "ARG_LABEL": "Customer ZIP", "ARG_SEQ": 4, "id": 44, "columnname": "CustomerZip", "Active": "Y" },
      { "ARG_ID": 45, "ARG_LABEL": "Customer State", "ARG_SEQ": 5, "id": 45, "columnname": "CustomerState", "Active": "Y" },
      { "ARG_ID": 46, "ARG_LABEL": "RO Close Date", "ARG_SEQ": 6, "id": 46, "columnname": "cdate", "Active": "Y" },
      // { "ARG_ID": 47, "ARG_LABEL": "No Grouping", "ARG_SEQ": 7, "id": 47, "columnname": "", "Active": "N" },
      { "ARG_ID": 53, "ARG_LABEL": "Source", "ARG_SEQ": 8, "id": 53, "columnname": "AP_source", "Active": "Y" }
      ];


  stores: any = []
  groupsArray: any = [];
  storename: any = ''
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  groupId: any = 0;
  storeIds: any = 0;


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



  reportOpenSub!: Subscription;
  reportGetting!: Subscription;
  Pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;

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
    console.log(this.selectedDataGrouping,'.........');
    
    this.shared.setTitle(this.comm.titleName + '-Parts Gross');
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
        if (this.shared.common.groupsandstores.length > 0) {
      this.groupsArray = this.shared.common.groupsandstores.filter((val: any) => val.sg_id != this.shared.common.reconID);
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
    this.setHeaderData()
    this.getPartsData();
  }
  setHeaderData() {
    const data = {
      title: 'Parts Gross',

      stores: this.storeIds,
      Department: this.Department,
      Paytype: this.Paytype,
      ToporBottom: this.TotalReport,
      fromdate: this.FromDate,
      todate: this.ToDate,
      groups: this.groups,
      servicetype: this.servicetype,
      GridView: this.GridView,
      partsSource: this.partsSource,
      count: 0,
    };
    this.shared.api.SetHeaderData({
      obj: data,
    });


  }

  initializeDates(type: any) {
    let dates: any = this.setdates.setDates(type)
    this.FromDate = dates[0];
    this.ToDate = dates[1];
    localStorage.setItem('time', type);
    this.DateType = type
    this.setDates(this.DateType)

  }
  getPartsData() {
    this.scrollCurrentposition = 0
    this.GridView = 'Global'
    this.NoData = false
    this.responcestatus = '';
    this.pageNumber = 0
    this.IndividualPartsGross = [];
    this.TotalPartsGross = [];
    this.setHeaderData();
    this.shared.spinner.show();
    if ((this.storeIds != '' && this.storeIds != '0') || this.selectedotherstoreids != '') {
      this.GetData();
      this.GetTotalData();
    } else {
      // this.NoData = true;
      this.shared.spinner.hide()
    }
  }
  GetData() {
    this.IndividualPartsGross = [];
    this.shared.spinner.show();
    const obj = {
      startdate: this.FromDate,
      enddate: this.ToDate,
      dealername: '',
      Advisorname: '',
      Store: this.selectedotherstoreids != undefined && this.selectedotherstoreids != '' && this.selectedotherstoreids != null ?
        (this.storeIds != '' ? this.storeIds.toString() + ',' + this.selectedotherstoreids.toString() : this.selectedotherstoreids.toString()) : this.storeIds.toString(),
      Labortype: this.Department.indexOf('Service') >= 0 ? this.servicetype.toString() : '',
      Saletype: this.Department.indexOf('Parts') >= 0 ? this.Paytype.toString() : '',
      SourceBulk: this.partsSource == 'B' ? 'Y' : '',
      SourceTire: this.partsSource == 'T' ? 'Y' : '',
      SourceWithout: this.partsSource == 'N' ? 'Y' : '',
      var1: this.selectedDataGrouping.length >= 1 ? this.selectedDataGrouping[0]?.columnname : '',
      var2: this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '',
      var3: '',
      RowType: 'D',
      UserID: this.comm.userId,
      PageNumber: this.pageNumber,
      PageSize: '100',
    };
    const curl = environment.apiUrl + this.comm.routeEndpoint + 'GetPartsGrossSummaryNewV1';
    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetPartsGrossSummaryNewV1', obj).subscribe(
      (res) => {
        const currentTitle = document.title;
        this.shared.api.logSaving(curl, {}, '', res.message, currentTitle);

        if (res.status == 200) {
          if (res.response != undefined) {
            if (res.response.length > 0) {
              this.IndividualPartsGross = [];
              this.IndividualPartsGross = res.response;
              this.responcestatus = this.responcestatus + 'I';
              let idi_len = this.IndividualPartsGross.length;
              this.IndividualPartsGross.some(function (x: any) {
                if (x.data2 != undefined && x.data2 != '') {
                  x.Data2 = JSON.parse(x.data2);
                  x.Dealer = '-';
                } else {
                  x.Dealer = '+';
                }
              });
              this.combineIndividualandTotal();
              // this.spinnerLoader=false;
            }
            else {
              // this.toast.show('Empty Response','');
              this.shared.spinner.hide();
              this.NoData = true;
              // this.spinnerLoader=false;
            }
          } else {
            // this.toast.show('Empty Response');
            this.shared.spinner.hide();
            this.NoData = true;
            // this.spinnerLoader=false;
          }
        } else {
          //this.toast.show('No Response', '');
          this.shared.spinner.hide();
          this.NoData = true;
          // this.spinnerLoader=false;
        }
      },
      (error) => {
        this.toast.show('502 Bad Gate Way Error', 'danger','Error');
        this.shared.spinner.hide();
        // this.spinnerLoader=false;
        this.combineIndividualandTotal()
      }
    );
  }
  GetTotalData() {
    this.TotalPartsGross = [];
    const obj = {
      startdate: this.FromDate,
      enddate: this.ToDate,
      dealername: '',
      Advisorname: '',
      Store: this.selectedotherstoreids != undefined && this.selectedotherstoreids != '' && this.selectedotherstoreids != null ?
        (this.storeIds != '' ? this.storeIds.toString() + ',' + this.selectedotherstoreids.toString() : this.selectedotherstoreids.toString()) : this.storeIds.toString(),
      Labortype: this.Department.indexOf('Service') >= 0 ? this.servicetype.toString() : '',
      Saletype: this.Department.indexOf('Parts') >= 0 ? this.Paytype.toString() : '',
      // PartsSource: this.partsSource=='All'? '': this.partsSource ,
      SourceBulk: this.partsSource == 'B' ? 'Y' : '',
      SourceTire: this.partsSource == 'T' ? 'Y' : '',
      SourceWithout: this.partsSource == 'N' ? 'Y' : '',
      var1: this.selectedDataGrouping[0].columnname == 'DealerName' ? '' : this.selectedDataGrouping[0].columnname,
      var2: this.selectedDataGrouping.length >= 2 ? (this.selectedDataGrouping[1]?.columnname == 'DealerName' ? '' : this.selectedDataGrouping[1]?.columnname) : '',
      var3: '',
      RowType: 'T',
      UserID: this.comm.userId,
      PageNumber: '',
      PageSize: '',
    };
    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetPartsGrossSummaryNewV1', obj).subscribe(
      (totalres) => {
        if (totalres.status == 200) {
          if (totalres.response != undefined) {
            if (totalres.response.length > 0) {
              this.TotalPartsGross = totalres.response.map((v: any) => ({
                ...v,
                Data2: [],
                DealerName: 'Reports Total',
                Dealer: '+',
              }));
              this.responcestatus = this.responcestatus + 'T';
              this.combineIndividualandTotal();
            } else {
              // this.toast.show('Empty Response','');
              this.shared.spinner.hide();
              this.NoData = true;
            }
          } else {
            // alert('Empty Response');
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
        //console.log(error);
        this.shared.spinner.hide();
        // this.NoData = true;
        this.combineIndividualandTotal()
      }
    );
  }
  spinnerLoader: boolean = false;
  combineIndividualandTotal() {
    if (this.responcestatus == 'IT' || this.responcestatus == 'TI') {
      if (this.TotalReport == 'B') {
        this.IndividualPartsGross.push(this.TotalPartsGross[0]);
        // this.PartsData = [...this.PartsData, ...this.IndividualPartsGross];
        this.PartsData = this.IndividualPartsGross;
      } else {
        this.IndividualPartsGross.unshift(this.TotalPartsGross[0]);
        // this.PartsData = [...this.PartsData, ...this.IndividualPartsGross];
        this.PartsData = this.IndividualPartsGross;
      }
      this.shared.spinner.hide();
      //console.log(this.PartsData);
      let position = this.scrollCurrentposition + 10
      setTimeout(() => {
        this.scrollcent.nativeElement.scrollTop = position
      }, 500);
    } else if (this.responcestatus == 'T') {
      // this.PartsData = [...this.PartsData, ...this.TotalPartsGross];
      this.PartsData = this.TotalPartsGross;
    } else if (this.responcestatus == 'I') {
      this.PartsData = [...this.PartsData, ...this.IndividualPartsGross];
      if (this.spinnerLoader == true) {
        this.shared.spinner.hide();
        this.spinnerLoader = false;
      }
    } else {
      if (this.PartsData.length == 0) {
        this.NoData = true;
      }
    }
  }
  GetServicedata() {
    this.scrollCurrentposition = 0
    this.GridView = 'Service'
    this.responcestatus = '';
    this.pageNumber = 0;
    this.NoData = false
    this.TotalServiceData = []
    this.IndividualServiceData = []
    this.shared.spinner.show();
    this.setHeaderData();
    this.GetServiceIndividualStoresData();
    this.GetTotalServiceData();
  }
  IndividualServiceData: any = []
  TotalServiceData: any = []
  ServcieGrossData: any = []
  GetServiceIndividualStoresData() {
    this.IndividualServiceData = [];
    this.shared.spinner.show();
    const obj = {
      startdate: this.FromDate,
      enddate: this.ToDate,
      dealername: '',
      Advisorname: '',
      Store: this.selectedotherstoreids != undefined && this.selectedotherstoreids != '' && this.selectedotherstoreids != null ?
        (this.storeIds != '' ? this.storeIds.toString() + ',' + this.selectedotherstoreids.toString() : this.selectedotherstoreids.toString()) : this.storeIds.toString(),
      Labortype: this.Department.indexOf('Service') >= 0 ? this.servicetype.toString() : '',
      Saletype: this.Department.indexOf('Parts') >= 0 ? this.Paytype.toString() : '',
      // PartsSource: this.partsSource=='All'? '': this.partsSource ,
      SourceBulk: this.partsSource == 'B' ? 'Y' : '',
      SourceTire: this.partsSource == 'T' ? 'Y' : '',
      SourceWithout: this.partsSource == 'N' ? 'Y' : '',
      // var1: this.Department.indexOf('Parts') >= 0 && this.Department.indexOf('Service') >= 0 ? this.path1.split(',')[0] : this.path1,
      // var2: this.Department.indexOf('Parts') >= 0 && this.Department.indexOf('Service') >= 0 ? this.path2.split(',')[0] : this.path2,
      // var3: this.path1 != 'DealerName' ? (this.Department.indexOf('Parts') >= 0 && this.Department.indexOf('Service') >= 0 ? this.path1.split(',')[1] : '') : this.path2.split(',')[1],
      var1: this.selectedDataGrouping.length >= 1 ? this.selectedDataGrouping[0]?.columnname : '',
      var2: this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '',
      var3: '',
      RowType: 'D',
      PageNumber: this.pageNumber,
      PageSize: '100',
    };
    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetPartsGrossSummaryNewServiceDD', obj).subscribe(
      (res) => {
        if (res.status == 200) {
          if (res.response != undefined) {
            if (res.response.length > 0) {
              this.IndividualServiceData = [];
              this.IndividualServiceData = res.response;
              this.responcestatus = this.responcestatus + 'I';
              let idi_len = this.IndividualServiceData.length;
              this.IndividualServiceData.some(function (x: any) {
                if (x.data2 != undefined && x.data2 != '') {
                  x.Data2 = JSON.parse(x.data2);
                  x.Dealer = '-';
                } else {
                  x.Dealer = '+';
                }
                // if (idi_len == 1) {
                //   x.Dealer = '-';
                // } else {
                //   x.Dealer = '+';
                // }
              });
              this.combineIndividualandTotalService();
            }
            else {
              // this.toast.show('Empty Response','');
              this.shared.spinner.hide();
              this.NoData = true;
            }
          } else {
            // this.toast.show('Empty Response');
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
        this.toast.show('502 Bad Gate Way Error', 'danger','Error');
        this.shared.spinner.hide();
        this.combineIndividualandTotalService()
      }
    );
  }
  GetTotalServiceData() {
    this.TotalServiceData = [];
    const obj = {
      startdate: this.FromDate,
      enddate: this.ToDate,
      dealername: '',
      Advisorname: '',
      Store: this.selectedotherstoreids != undefined && this.selectedotherstoreids != '' && this.selectedotherstoreids != null ?
        (this.storeIds != '' ? this.storeIds.toString() + ',' + this.selectedotherstoreids.toString() : this.selectedotherstoreids.toString()) : this.storeIds.toString(),
      Labortype: this.Department.indexOf('Service') >= 0 ? this.servicetype.toString() : '',
      Saletype: this.Department.indexOf('Parts') >= 0 ? this.Paytype.toString() : '',
      // PartsSource: this.partsSource=='All'? '': this.partsSource ,
      SourceBulk: this.partsSource == 'B' ? 'Y' : '',
      SourceTire: this.partsSource == 'T' ? 'Y' : '',
      SourceWithout: this.partsSource == 'N' ? 'Y' : '',
      // var1: this.Department.indexOf('Parts') >= 0 && this.Department.indexOf('Service') >= 0 ? this.path1.split(',')[0] : this.path1,
      // var2: this.Department.indexOf('Parts') >= 0 && this.Department.indexOf('Service') >= 0 ? this.path2.split(',')[0] : this.path2,
      // var3: this.path1 != 'DealerName' ? (this.Department.indexOf('Parts') >= 0 && this.Department.indexOf('Service') >= 0 ? this.path1.split(',')[1] : '') : this.path2.split(',')[1],
      var1: this.selectedDataGrouping[0].columnname == 'DealerName' ? '' : this.selectedDataGrouping[0].columnname,
      var2: this.selectedDataGrouping.length >= 2 ? (this.selectedDataGrouping[1]?.columnname == 'DealerName' ? '' : this.selectedDataGrouping[1]?.columnname) : '',
      var3: '',
      RowType: 'T',
      PageNumber: '',
      PageSize: '',
    };
    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetPartsGrossSummaryNewServiceDD', obj).subscribe(
      (totalres) => {
        if (totalres.status == 200) {
          if (totalres.response != undefined) {
            if (totalres.response.length > 0) {
              this.TotalServiceData = totalres.response.map((v: any) => ({
                ...v,
                Data2: [],
                DealerName: 'Reports Total',
                Dealer: '+',
              }));
              this.responcestatus = this.responcestatus + 'T';
              this.combineIndividualandTotalService();
            } else {
              // this.toast.show('Empty Response','');
              this.shared.spinner.hide();
              this.NoData = true;
            }
          } else {
            // alert('Empty Response');
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
        //console.log(error);
        this.shared.spinner.hide();
        // this.NoData = true;
        this.combineIndividualandTotalService()
      }
    );
  }
  combineIndividualandTotalService() {
    if (this.responcestatus == 'IT' || this.responcestatus == 'TI') {
      if (this.TotalReport == 'B') {
        this.IndividualServiceData.push(this.TotalServiceData[0]);
        this.ServcieGrossData = this.IndividualServiceData;
      } else {
        this.IndividualServiceData.unshift(this.TotalServiceData[0]);
        this.ServcieGrossData = this.IndividualServiceData;
      }
      this.shared.spinner.hide();
      //console.log(this.ServcieGrossData);
    } else if (this.responcestatus == 'T') {
      this.ServcieGrossData = this.TotalServiceData;
    } else if (this.responcestatus == 'I') {
      this.ServcieGrossData = [...this.ServcieGrossData, ...this.IndividualServiceData];
      if (this.spinnerLoader == true) {
        this.shared.spinner.hide();
        this.spinnerLoader = false;
      }
    } else {
      this.NoData = true;
    }
  }
  GetPartsDrilldata() {
    this.scrollCurrentposition = 0
    this.GridView = 'Parts'
    this.responcestatus = '';
    this.pageNumber = 0;
    this.NoData = false
    this.TotalPartsDrillData = []
    this.IndividualPartsDrillData = []
    this.shared.spinner.show();
    this.setHeaderData();
    this.GetPartsDrillIndividualStoresData();
    this.GetTotalPartsDrill();
  }
  IndividualPartsDrillData: any = []
  TotalPartsDrillData: any = []
  PartsDrillData: any = []
  GetPartsDrillIndividualStoresData() {
    this.IndividualPartsDrillData = [];
    this.shared.spinner.show();
    const obj = {
      startdate: this.FromDate,
      enddate: this.ToDate,
      dealername: '',
      Advisorname: '',
      Store: this.selectedotherstoreids != undefined && this.selectedotherstoreids != '' && this.selectedotherstoreids != null ?
        (this.storeIds != '' ? this.storeIds.toString() + ',' + this.selectedotherstoreids.toString() : this.selectedotherstoreids.toString()) : this.storeIds.toString(),
      Labortype: this.Department.indexOf('Service') >= 0 ? this.servicetype.toString() : '',
      Saletype: this.Department.indexOf('Parts') >= 0 ? this.Paytype.toString() : '',
      // PartsSource: this.partsSource=='All'? '': this.partsSource ,
      SourceBulk: this.partsSource == 'B' ? 'Y' : '',
      SourceTire: this.partsSource == 'T' ? 'Y' : '',
      SourceWithout: this.partsSource == 'N' ? 'Y' : '',

      var1: this.selectedDataGrouping.length >= 1 ? this.selectedDataGrouping[0]?.columnname : '',
      var2: this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '',
      var3: '',
      RowType: 'D',
      PageNumber: this.pageNumber,
      PageSize: '100',
    };
    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetPartsGrossSummaryNewPartsDD', obj).subscribe(
      (res) => {
        if (res.status == 200) {
          if (res.response != undefined) {
            if (res.response.length > 0) {
              this.IndividualPartsDrillData = [];
              this.IndividualPartsDrillData = res.response;
              this.responcestatus = this.responcestatus + 'I';
              let idi_len = this.IndividualPartsDrillData.length;
              this.IndividualPartsDrillData.some(function (x: any) {
                if (x.data2 != undefined && x.data2 != '') {
                  x.Data2 = JSON.parse(x.data2);
                  x.Dealer = '-';
                } else {
                  x.Dealer = '+';
                }
                // if (idi_len == 1) {
                //   x.Dealer = '-';
                // } else {
                //   x.Dealer = '+';
                // }
              });
              this.combineIndividualandTotalPartsDrill();
            }
            else {
              // this.toast.show('Empty Response','');
              this.shared.spinner.hide();
              this.NoData = true;
            }
          } else {
            // this.toast.show('Empty Response');
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
        this.toast.show('502 Bad Gate Way Error', 'danger','Error');
        this.shared.spinner.hide();
        this.combineIndividualandTotalPartsDrill()
      }
    );
  }
  GetTotalPartsDrill() {
    this.TotalPartsDrillData = [];
    const obj = {
      startdate: this.FromDate,
      enddate: this.ToDate,
      dealername: '',
      Advisorname: '',
      Store: this.selectedotherstoreids != undefined && this.selectedotherstoreids != '' && this.selectedotherstoreids != null ?
        (this.storeIds != '' ? this.storeIds.toString() + ',' + this.selectedotherstoreids.toString() : this.selectedotherstoreids.toString()) : this.storeIds.toString(),
      Labortype: this.Department.indexOf('Service') >= 0 ? this.servicetype.toString() : '',
      Saletype: this.Department.indexOf('Parts') >= 0 ? this.Paytype.toString() : '',
      // PartsSource: this.partsSource=='All'? '': this.partsSource ,
      SourceBulk: this.partsSource == 'B' ? 'Y' : '',
      SourceTire: this.partsSource == 'T' ? 'Y' : '',
      SourceWithout: this.partsSource == 'N' ? 'Y' : '',
      // var1: this.Department.indexOf('Parts') >= 0 && this.Department.indexOf('Service') >= 0 ? this.path1.split(',')[0] : this.path1,
      // var2: this.Department.indexOf('Parts') >= 0 && this.Department.indexOf('Service') >= 0 ? this.path2.split(',')[0] : this.path2,
      // var3: this.path1 != 'DealerName' ? (this.Department.indexOf('Parts') >= 0 && this.Department.indexOf('Service') >= 0 ? this.path1.split(',')[1] : '') : this.path2.split(',')[1],
      var1: this.selectedDataGrouping[0].columnname == 'DealerName' ? '' : this.selectedDataGrouping[0].columnname,
      var2: this.selectedDataGrouping.length >= 2 ? (this.selectedDataGrouping[1]?.columnname == 'DealerName' ? '' : this.selectedDataGrouping[1]?.columnname) : '',
      var3: '',
      RowType: 'T',
      PageNumber: '',
      PageSize: '',
    };
    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetPartsGrossSummaryNewPartsDD', obj).subscribe(
      (totalres) => {
        if (totalres.status == 200) {
          if (totalres.response != undefined) {
            if (totalres.response.length > 0) {
              this.TotalPartsDrillData = totalres.response.map((v: any) => ({
                ...v,
                Data2: [],
                DealerName: 'Reports Total',
                Dealer: '+',
              }));
              this.responcestatus = this.responcestatus + 'T';
              this.combineIndividualandTotalPartsDrill();
            } else {
              // this.toast.show('Empty Response','');
              this.shared.spinner.hide();
              this.NoData = true;
            }
          } else {
            // alert('Empty Response');
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
        //console.log(error);
        this.shared.spinner.hide();
        // this.NoData = true;
        this.combineIndividualandTotalPartsDrill()
      }
    );
  }
  combineIndividualandTotalPartsDrill() {
    if (this.responcestatus == 'IT' || this.responcestatus == 'TI') {
      if (this.TotalReport == 'B') {
        this.IndividualPartsDrillData.push(this.TotalPartsDrillData[0]);
        this.PartsDrillData = this.IndividualPartsDrillData;
      } else {
        this.IndividualPartsDrillData.unshift(this.TotalPartsDrillData[0]);
        this.PartsDrillData = this.IndividualPartsDrillData;
      }
      this.shared.spinner.hide();
      //console.log(this.PartsDrillData);
    } else if (this.responcestatus == 'T') {
      this.PartsDrillData = this.TotalPartsDrillData;
    } else if (this.responcestatus == 'I') {
      this.PartsDrillData = [...this.PartsDrillData, ...this.IndividualPartsDrillData];
      if (this.spinnerLoader == true) {
        this.shared.spinner.hide();
        this.spinnerLoader = false;
      }
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
  expandorcollapse(ind: any, e: any, ref: any, Item: any, parentData: any) {
    let id = (e.target as Element).id;
    if ((this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '') == '') {
      this.openDetails(Item, parentData, '1');
    }
    if (id == 'D_' + ind) {
      if (ref == '-') {
        Item.Dealer = '+';
      }
      if (ref == '+') {
        Item.Dealer = '-';
      }
    }
  }
  openDetails(Item: any, ParentItem: any, ref: any) {
    if (ref == '2') {
      if (Item.data2 != undefined) {
        const DetailsPartsPerson = this.shared.ngbmodal.open(PartsGrossDetails, {
          backdrop: 'static',
        });
        DetailsPartsPerson.componentInstance.Partsdetails = [
          {
            StartDate: this.FromDate,
            EndDate: this.ToDate,
            var1: this.selectedDataGrouping.length >= 1 ? this.selectedDataGrouping[0]?.columnname : '',
            var2: this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '',
            var3: '',
            var1Value: Item.data1,
            var2Value: Item.data2,
            var3Value: '',
            userName: Item.data2,
            type: Item.type,
            Labortype: this.Department.indexOf('Service') >= 0 ? this.servicetype.toString() : '',
            Saletype: this.Department.indexOf('Parts') >= 0 ? this.Paytype.toString() : '',
            Store: this.selectedDataGrouping[0].columnname  == 'DealerName' ? ParentItem.DealerID : Item.DealerID,
            PartsSource: this.partsSource == 'All' ? '' : this.partsSource,
          },
        ];
      }
    }
    if (ref == '1') {
      if (Item.data1 != undefined && Item.data1 != 'Reports Total') {
        const DetailsPartsPerson = this.shared.ngbmodal.open(PartsGrossDetails, {
          // size:'xl',
          backdrop: 'static',
        });
        DetailsPartsPerson.componentInstance.Partsdetails = [
          {
            StartDate: this.FromDate,
            EndDate: this.ToDate,
            var1: this.selectedDataGrouping.length >= 1 ? this.selectedDataGrouping[0]?.columnname : '',
            var2: this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '',
            var3: '',
            var1Value: Item.data1,
            var2Value: Item.data2 == undefined ? '' : Item.data2,
            var3Value: '',
            userName: Item.data1,
            type: Item.type,
            Labortype: this.Department.indexOf('Service') >= 0 ? this.servicetype.toString() : '',
            Saletype: this.Department.indexOf('Parts') >= 0 ? this.Paytype.toString() : '',
            Store: this.selectedDataGrouping[0].columnname  == 'DealerName' ? Item.DealerID : '',
            PartsSource: this.partsSource == 'All' ? '' : this.partsSource,
          },
        ];
        DetailsPartsPerson.result.then(
          (data) => { },
          (reason) => {
            // on dismiss
          }
        );
      }
    }
  }
  isDesc: boolean = false;
  column: string = 'CategoryName';
  sort(property: any) {
    this.isDesc = !this.isDesc; //change the direction
    this.column = property;
    let direction = this.isDesc ? 1 : -1;
    this.PartsData.sort(function (a: any, b: any) {
      if (a[property] < b[property]) {
        return -1 * direction;
      } else if (a[property] > b[property]) {
        return 1 * direction;
      } else {
        return 0;
      }
    });
  }
  ngAfterViewInit(): void {
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.comm.pageName == 'Parts Gross') {
        if (res.obj.storesData != undefined) {
          this.groupsArray = res.obj.storesData;
          this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
          this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_name : this.groupName = ''
          this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
          this.getStoresandGroupsValues()
        }
      }
    })

    this.excel = this.shared.api.getExportToExcelAllReports().subscribe((res) => {
      if (this.excel != undefined) {
        if (res.obj.title == 'Parts Gross') {
          if (res.obj.state == true) {
            this.exportToExcel();
          }
        }
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res) => {
      if (this.print != undefined) {
        if (res.obj.title == 'Parts Gross') {
          if (res.obj.statePrint == true) {
            //     this.GetPrintData();
          }
        }
      }
    });
    this.Pdf = this.shared.api.getExportToPDFAllReports().subscribe((res) => {
      if (this.Pdf != undefined) {
        if (res.obj.title == 'Parts Gross') {
          if (res.obj.statePDF == true) {
            //      this.generatePDF();
          }
        }
      }
    });
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res) => {
      if (this.email != undefined) {
        if (res.obj.title == 'Parts Gross') {
          if (res.obj.stateEmailPdf == true) {
            //    this.sendEmailData(res.obj.Email, res.obj.notes, res.obj.from);
          }
        }
      }
    });
  }
  ngOnDestroy() {
    // this.reportOpenSub.unsubscribe()
    // this.reportGetting.unsubscribe()
    // this.Pdf.unsubscribe()
    // this.print.unsubscribe()
    // this.email.unsubscribe()
    // this.excel.unsubscribe()
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
  StoresData(data: any) {
    this.storeIds = data.storeids;
    this.groupId = data.groupId;
    this.storename = data.storename;
    this.groupName = data.groupName;
    this.storecount = data.storecount;
    this.storedisplayname = data.storedisplayname;
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
  SelectedData(val: any) {
    const index = this.selectedDataGrouping.findIndex((i: any) => i == val);
    if (index >= 0) {
      this.selectedDataGrouping.splice(index, 1);
    } else {
      if (this.selectedDataGrouping.length >= 2) {
        this.toast.show('Select up to 2 Filters only to Group your data', 'warning', 'Warning');
      } else {
        this.selectedDataGrouping.push(val);
      }
    }
  }

  pageNumber: any = 0;
  Scrollpercent: any = 0;
  scrollCurrentposition: any = 0
  @ViewChild('scrollcent') scrollcent!: ElementRef;
  updateVerticalScroll(event: any, block: any): void {
    this.scrollCurrentposition = event.target.scrollTop
    const scrollDemo = document.querySelector('#scrollcent') as HTMLElement;
    this.Scrollpercent = Math.round(
      (event.target.scrollTop /
        (event.target.scrollHeight - scrollDemo.clientHeight)) *
      100
    );
    if (
      event.target.scrollTop + event.target.clientHeight >=
      event.target.scrollHeight - 2
    ) {
      if (block == 'Global') {
        if (this.IndividualPartsGross.length > 99) {
          this.pageNumber++
          this.responcestatus = '';
          this.spinnerLoader = true
          this.GetData()
        }
      }
      else if (block == 'Service') {
        if (this.IndividualServiceData.length > 99) {
          this.pageNumber++
          this.responcestatus = '';
          this.spinnerLoader = true
          this.GetServiceIndividualStoresData()
        }
      }
      else if (block == 'Parts') {
        if (this.IndividualPartsDrillData.length > 99) {
          this.pageNumber++
          this.responcestatus = '';
          this.spinnerLoader = true
          this.GetPartsDrillIndividualStoresData()
        }
      }
    }
  }
  activePopover: number = -1;

  togglePopover(popoverIndex: number) {
    this.activePopover = this.activePopover === popoverIndex ? -1 : popoverIndex;
  }
  saleandservice: any = ['R', 'W', 'C', 'T', 'I', 'E']

  multipleorsingle(block: any, e: any) {

    if (block == 'Dept') {
      const index = this.Department.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.Department.splice(index, 1);
        if (this.Department.indexOf('Service') >= 0 && this.Department.indexOf('Parts') >= 0) {
          this.Paytype = ['R', 'W']
          this.servicetype = ['C', 'T', 'I', 'E']
          this.saleandservice = []
          this.saleandservice = [...this.Paytype, ...this.servicetype]

        }
        else if (this.Department.indexOf('Service') >= 0) {
          this.servicetype = ['C', 'T', 'I', 'E']
          this.Paytype = []
          this.saleandservice = []
          this.saleandservice = [...this.Paytype, ...this.servicetype]

        }
        else if (this.Department.indexOf('Parts') >= 0) {
          this.Paytype = ['R', 'W']
          this.servicetype = []
          this.saleandservice = []
          this.saleandservice = [...this.Paytype, ...this.servicetype]

        } else {
          this.Paytype = []
          this.servicetype = []
          this.saleandservice = []
          this.saleandservice = [...this.Paytype, ...this.servicetype]
        }
      } else {
        this.Department.push(e);
        if (this.Department.indexOf('Service') >= 0 && this.Department.indexOf('Parts') >= 0) {
          this.Paytype = ['R', 'W']
          this.servicetype = ['C', 'T', 'I', 'E']
          this.saleandservice = []
          this.saleandservice = [...this.Paytype, ...this.servicetype]

        }
        else if (this.Department.indexOf('Service') >= 0) {
          this.servicetype = ['C', 'T', 'I', 'E']
          this.Paytype = []
          this.saleandservice = []
          this.saleandservice = [...this.Paytype, ...this.servicetype]

        }
        else if (this.Department.indexOf('Parts') >= 0) {
          this.Paytype = ['R', 'W']
          this.servicetype = []
          this.saleandservice = []
          this.saleandservice = [...this.Paytype, ...this.servicetype]

        }
      }
      console.log(this.saleandservice);

    }
    if (block == 'PT') {
      const index = this.Paytype.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.Paytype.splice(index, 1);
        this.saleandservice = []
        this.saleandservice = [...this.Paytype, ...this.servicetype]
      } else {
        this.Paytype.push(e);
        this.saleandservice = []
        this.saleandservice = [...this.Paytype, ...this.servicetype]
      }
    }
    if (block == 'PS') {
      this.partsSource = e
    }
    if (block == 'LT') {
      const index = this.servicetype.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.servicetype.splice(index, 1);
        this.saleandservice = []
        this.saleandservice = [...this.Paytype, ...this.servicetype]
      } else {
        this.servicetype.push(e);
        this.saleandservice = []
        this.saleandservice = [...this.Paytype, ...this.servicetype]
      }
    }

    if (block == 'TB') {
      this.TotalReport = e;

    }

  }
  viewreport() {

    this.activePopover = -1
    if (this.selectedDataGrouping.length == 0 || (this.storeIds.length == 0 && this.selectedotherstoreids.length == 0) || this.DateType.length == 0 || this.DateType == undefined) {
      if (this.selectedDataGrouping.length == 0) {
        this.toast.show('Please select atleast one Grouping', 'warning', 'Warning');
      }
      if (this.storeIds.length == 0 && this.selectedotherstoreids.length == 0) {
        this.toast.show('Please select atleast one Store', 'warning', 'Warning');
      }
      if (this.DateType.length == 0 || this.DateType == undefined) {
        this.toast.show('Please select atleast one Time Frame', 'warning', 'Warning');
      }

    } else {
      console.log(this.selectedDataGrouping,'........... Values');
      
      this.setHeaderData();
      if (this.GridView == 'Global') {
        this.getPartsData();
      } else if (this.GridView == 'Service') {
        this.GetServicedata();
      } else if (this.GridView == 'Parts') {
        this.GetPartsDrilldata();
      }
    }
  }
  ExcelStoreNames: any = [];
  exportToExcel() {
    let storeNames: any = [];
    let store = this.storeIds
    storeNames = this.comm.groupsandstores.filter((v: any) => v.sg_id == this.groups)[0].Stores.filter((item: any) =>
      store.some((cat: any) => cat === item.ID.toString())
    );
    if (store.length == this.comm.groupsandstores.filter((v: any) => v.sg_id == this.groups)[0].Stores.length) {
      this.ExcelStoreNames = 'All Stores'
    } else {
      this.ExcelStoreNames = storeNames.map(function (a: any) {
        return a.storename;
      });
    }
    const PartsData = this.PartsData.map((_arrayElement: any) =>
      Object.assign({}, _arrayElement)
    );
    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('Parts Gross');
    worksheet.views = [
      {
        state: 'frozen',
        ySplit: 18, // Number of rows to freeze (2 means the first two rows are frozen)
        topLeftCell: 'A19', // Specify the cell to start freezing from (in this case, the third row)
        showGridLines: false,
      },
    ];
    worksheet.addRow('');
    const titleRow = worksheet.addRow(['Parts Gross']);
    titleRow.eachCell((cell: any, number: any) => {
      cell.alignment = { indent: 1, vertical: 'middle', horizontal: 'left' };
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
    groups.value = this.comm.groupsandstores.filter((val: any) => val.sg_id == this.groups.toString())[0].sg_name;
    groups.font = { name: 'Arial', family: 4, size: 9 };
    const Brands = worksheet.getCell('B9');
    Brands.value = 'Brands :';
    const brands = worksheet.getCell('D9');
    brands.value = '-';
    brands.font = { name: 'Arial', family: 4, size: 9 };
    const Stores1 = worksheet.getCell('B10');
    Stores1.value = 'Stores :';
    worksheet.mergeCells('D10', 'O12');
    const stores1 = worksheet.getCell('D10');
    stores1.value =
      this.ExcelStoreNames == 0
        ? 'All Stores'
        : this.ExcelStoreNames == null
          ? '-'
          : this.ExcelStoreNames.toString().replaceAll(',', ', ');
    stores1.font = { name: 'Arial', family: 4, size: 9 };
    stores1.alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    const Filters = worksheet.addRow(['Filters :']);
    Filters.getCell(1).font = { name: 'Arial', family: 4, size: 9, bold: true };
    const DealType = worksheet.getCell('B13');
    DealType.value = 'Department :';
    const dealtype = worksheet.getCell('D13');
    dealtype.value = this.Department == '' ? '-' : this.Department == null ? '-' : this.Department.toString().replaceAll(',', ', ');
    dealtype.font = { name: 'Arial', family: 4, size: 9 };
    const DealStatus = worksheet.getCell('B14');
    DealStatus.value = 'Sale Type :';
    const dealstatus = worksheet.getCell('D14');
    dealstatus.value =
      this.Paytype == '' ? '-' : this.Paytype == null ? '-' : this.Paytype.toString().replaceAll(',', ', ');
    dealstatus.font = { name: 'Arial', family: 4, size: 9 };
    const Source = worksheet.getCell('B15');
    Source.value = 'Source :';
    const source = worksheet.getCell('D15');
    source.value = '-';
    source.font = { name: 'Arial', family: 4, size: 9 };
    worksheet.addRow('');
    let dateYear = worksheet.getCell('A17');
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
    dateYear.border = { right: { style: 'thin' } };
    worksheet.mergeCells('B17', 'G17');
    let totalparts = worksheet.getCell('B17');
    totalparts.value = 'Total Parts';
    totalparts.alignment = { vertical: 'middle', horizontal: 'center' };
    totalparts.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
      color: { argb: 'FFFFFF' },
    };
    totalparts.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '2a91f0' },
      bgColor: { argb: 'FF0000FF' },
    };
    totalparts.border = { right: { style: 'thin' } };
    worksheet.mergeCells('H17', 'K17');
    let mechanical = worksheet.getCell('H17');
    mechanical.value = 'Mechanical';
    mechanical.alignment = { vertical: 'middle', horizontal: 'center' };
    mechanical.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
      color: { argb: 'FFFFFF' },
    };
    mechanical.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '2a91f0' },
      bgColor: { argb: 'FF0000FF' },
    };
    mechanical.border = { right: { style: 'thin' } };
    worksheet.mergeCells('L17', 'O17');
    let retail = worksheet.getCell('L17');
    retail.value = 'Retail/Wholesale';
    retail.alignment = { vertical: 'middle', horizontal: 'center' };
    retail.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
      color: { argb: 'FFFFFF' },
    };
    retail.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '2a91f0' },
      bgColor: { argb: 'FF0000FF' },
    };
    retail.border = { right: { style: 'thin' } };
    worksheet.mergeCells('P17', 'R17');
    let performance = worksheet.getCell('P17');
    performance.value = 'Performance';
    performance.alignment = { vertical: 'middle', horizontal: 'center' };
    performance.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
      color: { argb: 'FFFFFF' },
    };
    performance.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '2a91f0' },
      bgColor: { argb: 'FF0000FF' },
    };
    performance.border = { right: { style: 'thin' } };
    let Headings = [
      FromDate + ' - ' + ToDate + ', ' + PresentYear,
      'Sales',
      // 'Discounts',
      'Gross',
      'Pace',
      'Target',
      'Diff',
      'Gross',
      'Pace',
      'Target',
      'Diff',
      'Gross',
      'Pace',
      'Target',
      'Diff',
      'Parts/RO',
      'Lost/day',
      'GP%',
    ];
    const headerRow = worksheet.addRow(Headings);
    headerRow.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: false,
      color: { argb: 'FFFFFF' },
    };
    headerRow.alignment = { indent: 1, vertical: 'middle', horizontal: 'center' };
    headerRow.eachCell((cell: any, number: any) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '788494' },
        bgColor: { argb: 'FF0000FF' },
      };
      cell.border = { right: { style: 'thin' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });
    for (const d of PartsData) {
      const Data1 = worksheet.addRow([
        d.data1 == '' ? '-' : d.data1 == null ? '-' : (this.selectedDataGrouping[0]?.columnname == 'cdate' && d.data1 != 'Reports Total') ? this.shared.datePipe.transform(d.data1, 'MM/dd/YYYY') : d.data1,
        d.Total_PartsSale == '' ? '-' : d.Total_PartsSale == null ? '-' : d.Total_PartsSale,
        // '-',
        d.Total_PartsGross == '' ? '-' : d.Total_PartsGross == null ? '-' : d.Total_PartsGross,
        d.Total_PartsGross_Pace == '' ? '-' : d.Total_PartsGross_Pace == null ? '-' : d.Total_PartsGross_Pace,
        d.Total_PartsGrossTarget == '' ? '-' : d.Total_PartsGrossTarget == null ? '-' : d.Total_PartsGrossTarget,
        d.Total_PartsGross_Diff == '' ? '-' : d.Total_PartsGross_Diff == null ? '-' : d.Total_PartsGross_Diff,
        d.ServiceGross == '' ? '-' : d.ServiceGross == null ? '-' : d.ServiceGross,
        d.ServiceGross_Pace == '' ? '-' : d.ServiceGross_Pace == null ? '-' : d.ServiceGross_Pace,
        d.ServiceGross_Target == '' ? '-' : d.ServiceGross_Target == null ? '-' : d.ServiceGross_Target,
        d.ServiceGross_Diff == '' ? '-' : d.ServiceGross_Diff == null ? '-' : d.ServiceGross_Diff,
        d.PartsGross == '' ? '-' : d.PartsGross == null ? '-' : d.PartsGross,
        d.PartsGross_Pace == '' ? '-' : d.PartsGross_Pace == null ? '-' : d.PartsGross_Pace,
        d.PartsGross_Target == '' ? '-' : d.PartsGross_Target == null ? '-' : d.PartsGross_Target,
        d.PartsGross_Diff == '' ? '-' : d.PartsGross_Diff == null ? '-' : d.PartsGross_Diff,
        d.Parts_RO == '' ? '-' : d.Parts_RO == null ? '-' : d.Parts_RO,
        d.Lost_PerDay == '' ? '-' : d.Lost_PerDay == null ? '-' : d.Lost_PerDay,
        d.Retention == '' ? '-' : d.Retention == null ? '-' : d.Retention + '%',
      ]);
      // Data1.outlineLevel = 1; // Grouping level 1
      Data1.font = { name: 'Arial', family: 4, size: 9 };
      Data1.getCell(1).alignment = {
        indent: 1,
        vertical: 'middle',
        horizontal: 'left',
      };
      Data1.eachCell((cell: any, number: any) => {
        cell.border = { right: { style: 'thin' } };
        if (number == 15) {
          cell.numFmt = '#,##0.00';
        } else {
          cell.numFmt = '$#,##0';
        }
        if (number != 1) {
          cell.alignment = { vertical: 'middle', horizontal: 'center', indent: 1 };
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
      if (d.Total_PartsSale < 0) {
        Data1.getCell(2).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
      } if (d.Total_PartsGross < 0) {
        Data1.getCell(4).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
      } if (d.Total_PartsGross_Pace < 0) {
        Data1.getCell(5).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
      } if (d.Total_PartsGrossTarget < 0) {
        Data1.getCell(6).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
      } if (d.Total_PartsGross_Diff < 0) {
        Data1.getCell(7).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
      } if (d.ServiceGross < 0) {
        Data1.getCell(8).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
      } if (d.ServiceGross_Pace < 0) {
        Data1.getCell(9).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
      } if (d.ServiceGross_Target < 0) {
        Data1.getCell(10).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
      } if (d.Total_PartsGross_Diff < 0) {
        Data1.getCell(11).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
      } if (d.PartsGross < 0) {
        Data1.getCell(12).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
      } if (d.PartsGross_Pace < 0) {
        Data1.getCell(13).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
      } if (d.PartsGross_Target < 0) {
        Data1.getCell(14).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
      } if (d.PartsGross_Diff < 0) {
        Data1.getCell(15).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
      } if (d.Parts_RO < 0) {
        Data1.getCell(16).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
      } if (d.Lost_PerDay < 0) {
        Data1.getCell(17).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
      } if (d.Retention < 0) {
        Data1.getCell(18).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
      }
      if (d.Data2 != undefined) {
        for (const d1 of d.Data2) {
          const Data2 = worksheet.addRow([
            d1.data2 == '' ? '-' : d1.data2 == null ? '-' : this.selectedDataGrouping.length==2 && this.selectedDataGrouping[0]?.columnname == 'cdate' ? this.shared.datePipe.transform(d1.data2, 'MM/dd/YYYY') : d1.data2,
            d1.Total_PartsSale == '' ? '-' : d1.Total_PartsSale == null ? '-' : d1.Total_PartsSale,
            d1.Total_PartsGross == '' ? '-' : d1.Total_PartsGross == null ? '-' : d1.Total_PartsGross,
            d1.Total_PartsGross_Pace == '' ? '-' : d1.Total_PartsGross_Pace == null ? '-' : d1.Total_PartsGross_Pace,
            '-',
            '-',
            d1.ServiceGross == '' ? '-' : d1.ServiceGross == null ? '-' : d1.ServiceGross,
            d1.ServiceGross_Pace == '' ? '-' : d1.ServiceGross_Pace == null ? '-' : d1.ServiceGross_Pace,
            d1.ServiceGross_Target == '' ? '-' : d1.ServiceGross_Target == null ? '-' : d1.ServiceGross_Target,
            d1.ServiceGross_Diff == '' ? '-' : d1.ServiceGross_Diff == null ? '-' : d1.ServiceGross_Diff,
            d1.PartsGross == '' ? '-' : d1.PartsGross == null ? '-' : d1.PartsGross,
            d1.PartsGross_Pace == '' ? '-' : d1.PartsGross_Pace == null ? '-' : d1.PartsGross_Pace,
            d1.PartsGross_Target == '' ? '-' : d1.PartsGross_Target == null ? '-' : d1.PartsGross_Target,
            d1.PartsGross_Diff == '' ? '-' : d1.PartsGross_Diff == null ? '-' : d1.PartsGross_Diff,
            d1.Parts_RO == '' ? '-' : d1.Parts_RO == null ? '-' : d1.Parts_RO,
            d1.Lost_PerDay == '' ? '-' : d1.Lost_PerDay == null ? '-' : d1.Lost_PerDay,
            d1.Retention == '' ? '-' : d1.Retention == null ? '-' : d1.Retention + ' %',
          ]);
          Data2.outlineLevel = 1; // Grouping level 2
          Data2.font = { name: 'Arial', family: 4, size: 8 };
          Data2.getCell(1).alignment = {
            indent: 2,
            vertical: 'middle',
            horizontal: 'left',
          };
          Data2.eachCell((cell: any, number: any) => {
            cell.border = { right: { style: 'thin' } };
            if (number == 15) {
              cell.numFmt = '#,##0.00';
            } else {
              cell.numFmt = '$#,##0';
            }
            if (number != 1) {
              cell.alignment = {
                vertical: 'middle',
                horizontal: 'center',
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
          if (d1.Total_PartsSale < 0) {
            Data2.getCell(2).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
          } if (d1.Total_PartsGross < 0) {
            Data2.getCell(4).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
          } if (d1.Total_PartsGross_Pace < 0) {
            Data2.getCell(5).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
          } if (d1.Total_PartsGrossTarget < 0) {
            Data2.getCell(6).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
          } if (d1.Total_PartsGross_Diff < 0) {
            Data2.getCell(7).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
          } if (d1.ServiceGross < 0) {
            Data2.getCell(8).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
          } if (d1.ServiceGross_Pace < 0) {
            Data2.getCell(9).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
          } if (d1.ServiceGross_Target < 0) {
            Data2.getCell(10).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
          } if (d1.Total_PartsGross_Diff < 0) {
            Data2.getCell(11).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
          } if (d1.PartsGross < 0) {
            Data2.getCell(12).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
          } if (d1.PartsGross_Pace < 0) {
            Data2.getCell(13).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
          } if (d1.PartsGross_Target < 0) {
            Data2.getCell(14).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
          } if (d1.PartsGross_Diff < 0) {
            Data2.getCell(15).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
          } if (d1.Parts_RO < 0) {
            Data2.getCell(16).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
          } if (d1.Lost_PerDay < 0) {
            Data2.getCell(17).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
          } if (d1.Retention < 0) {
            Data2.getCell(18).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
          }
        }
      }
      if (d.data1 === 'Reports Total') {
        Data1.eachCell((cell) => {
          cell.font = { name: 'Arial', family: 4, size: 9, bold: true };
          // cell.border = {
          //   top: { style: 'thin' },
          //   bottom: { style: 'thin' },
          // };
        });
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
    worksheet.addRow([]);

    this.shared.exportToExcel(workbook, 'Parts Gross_' + DATE_EXTENSION);

  }

}