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
import { NgbModalModule, NgbModalRef } from '@ng-bootstrap/ng-bootstrap';

@Component({
  selector: 'app-dashboard',
  imports: [SharedModule, BsDatepickerModule, DateRangePicker, Stores,NgbModalModule],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
})
export class Dashboard {
  ServiceData: any = [];
  IndividualServiceGross: any = [];
  TotalServiceGross: any = [];


  TotalReport: string = 'T';
  NoData: boolean = false;
  Paytype: any = ['Customerpay_0', 'Warranty_1', 'Internal_2'];
  Department: any = ['Parts'];
  saletype: any = []

  responcestatus: any = '';
  groups: any = 1;
  lastyearDate: any;

  reportOpenSub!: Subscription;
  reportGetting!: Subscription;
  Pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;
  otherstoreid: any = '';
  selectedotherstoreids: any = '';

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

  selectedDataGrouping: any = [];
  dataGrouping: any = [
    { "ARG_ID": 62, "ARG_LABEL": "Store", "ARG_SEQ": 0, "id": 62, "columnname": "Store_Name", "Active": "Y" },
    { "ARG_ID": 63, "ARG_LABEL": "Sale Type", "ARG_SEQ": 2, "id": 63, "columnname": "ASG_Subtype_Detail", "Active": "Y" }
  ]

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
    this.getSaleTypeList();

    if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
      this.groupId = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Preferences
      this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids.split(',')
    }
    if (this.shared.common.groupsandstores.length > 0) {
      this.groupsArray = this.shared.common.groupsandstores.filter((val: any) => val.sg_id != this.shared.common.reconID);
      this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
      this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_Name : this.groupName = ''
      this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
      this.getStoresandGroupsValues()
    }

    this.shared.setTitle(this.comm.titleName + '-Parts Gross GL');
    let today = new Date();
    if (today.getDate() == 1) {
      this.initializeDates('LMGL')
    } else if (today.getDate() > 1 && today.getDate() < 5) {
      this.initializeDates('LMGL')
    } else {
      this.initializeDates('MTD')
    }
    this.setHeaderData();
    this.getServiceData();
  }
  asoftime: any = [];
  ngOnInit() {
    const obj = {};
    this.shared.api
      .postmethod(this.comm.routeEndpoint + 'GetSetviceGLASofTime', obj)
      .subscribe((res: any) => {
        //console.log(res.response);
        if (res.status == 200) {
          if (res.response != undefined) this.asoftime = res.response;
        }
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


  setHeaderData() {
    const data = {
      title: 'Parts Gross GL',
      dataGroupings: this.selectedDataGrouping,
      Department: this.Department,
      Paytype: this.Paytype,
      ToporBottom: this.TotalReport,
      fromdate: this.FromDate,
      todate: this.ToDate,
      groups: this.groupId,
      otherstoreids: this.otherstoreid, selectedotherstoreids: this.selectedotherstoreids
    };
    this.shared.api.SetHeaderData({
      obj: data,
    });

  }
  getServiceData() {
    if (this.storeIds != '' || this.selectedotherstoreids != '') {
      this.responcestatus = '';
      this.shared.spinner.show();
      this.GetData();
      this.GetTotalData();
    } else {
      // this.NoData = true;
    }
  }
  GetData() {
    this.IndividualServiceGross = [];
    const obj = {
      startdate: this.FromDate.replaceAll('/', '-'),
      enddate: this.ToDate.replaceAll('/', '-'),
      StoreID: this.selectedotherstoreids != undefined && this.selectedotherstoreids != '' && this.selectedotherstoreids != null ?
        (this.storeIds != '' ? this.storeIds + ',' + this.selectedotherstoreids.toString() : this.selectedotherstoreids.toString()) : this.storeIds.toString(),
      PaytypeC: this.Paytype[0] == 'Customerpay_0' ? 'C' : '',
      PaytypeW: this.Paytype[1] == 'Warranty_1' ? 'W' : '',
      PaytypeI: this.Paytype[2] == 'Internal_2' ? 'I' : '',

      RO_CloseDate: '',
      var1: this.selectedDataGrouping.length >= 1 ? this.selectedDataGrouping[0]?.columnname : '',
      var2: this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '',
      var3: this.selectedDataGrouping.length == 3 ? this.selectedDataGrouping[2]?.columnname : '',
      type: 'D',
      Details: this.Department.indexOf('Details') >= 0 ? 'D' : '',
      saletype: this.saletype.length == this.partsSaleType.length ? '' : this.saletype.toString()
    };
    const curl = environment.apiUrl + this.comm.routeEndpoint + 'GetPartsSummaryGL';
    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetPartsSummaryGL', obj).subscribe(
      (res) => {
        const currentTitle = document.title;
        this.shared.api.logSaving(curl, {}, '', res.message, currentTitle);
        if (res.status == 200) {
          if (res.response != undefined) {
            if (res.response.length > 0) {
              this.IndividualServiceGross = res.response;
              this.responcestatus = this.responcestatus + 'I';
              this.NoData = false;
              let path2 = this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '';
              let path3 = this.selectedDataGrouping.length == 3 ? this.selectedDataGrouping[2]?.columnname : '';
              console.log(path3, 'Path 3');

              this.IndividualServiceGross.some(function (x: any) {
                if (x.Details != undefined) {
                  x.Data2 = JSON.parse(x.Details);
                  if (path3 != '') {
                    let data = JSON.parse(x.Details);
                    x.Data2 = data.reduce(
                      (r: any, { data2, data3, StoreID }: any) => {
                        if (!r.some((o: any) => o.data2 == data2)) {
                          r.push({
                            data2,
                            data3,
                            StoreID,
                            data2sign: '-',
                            subdata: data.filter((v: any) => v.data2 == data2),
                          });
                        }
                        return r;
                      },
                      []
                    );
                  }
                }
                if (path2 == '') {
                  x.Dealer = '+';
                } else {
                  x.Dealer = '-';
                }
              });
              this.combineIndividualandTotal();
            } else {
              this.shared.spinner.hide();
              this.NoData = true;
            }
          } else {
            this.shared.spinner.hide();
            this.NoData = true;
          }
        } else {
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
      startdate: this.FromDate.replaceAll('/', '-'),
      enddate: this.ToDate.replaceAll('/', '-'),
      StoreID: this.selectedotherstoreids != undefined && this.selectedotherstoreids != '' && this.selectedotherstoreids != null ?
        (this.storeIds != '' ? this.storeIds + ',' + this.selectedotherstoreids.toString() : this.selectedotherstoreids.toString()) : this.storeIds.toString(),
      PaytypeC: this.Paytype[0] == 'Customerpay_0' ? 'C' : '',
      PaytypeW: this.Paytype[1] == 'Warranty_1' ? 'W' : '',
      PaytypeI: this.Paytype[2] == 'Internal_2' ? 'I' : '',

      RO_CloseDate: '',
      var1: 'Store_Name',
      var2: this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '',
      var3: this.selectedDataGrouping.length == 3 ? this.selectedDataGrouping[2]?.columnname : '',
      type: 'T',
      Details: this.Department.indexOf('Details') >= 0 ? 'D' : '',
      saletype: this.saletype.length == this.partsSaleType.length ? '' : this.saletype.toString()
    };
    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetPartsSummaryGL', obj).subscribe(
      (totalres) => {
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
              this.TotalServiceGross.some(function (x: any) {
                if (x.Details != undefined) {
                  x.Data2 = JSON.parse(x.Details);
                  x.Dealer = '-';
                }
              });
              this.combineIndividualandTotal();
            } else {

              this.shared.spinner.hide();
              // this.NoData = true;
            }
          } else {

            this.shared.spinner.hide();
            // this.NoData = true;
          }
        } else {

          this.shared.spinner.hide();
          // this.NoData = true;
        }
      },
      (error) => {
        this.toast.show('502 Bad Gate Way Error', 'danger', 'Error');
        this.shared.spinner.hide();
        // this.NoData = true;
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
      console.log(this.ServiceData, 'Service Data');

      this.shared.spinner.hide();
      //console.log(this.ServiceData);
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
    } else if (value < 0) {
      return false;
    }
    return true;
  }
  expandorcollapse(ind: any, e: any, ref: any, Item: any, parentData: any, tmp: any) {
    let id = (e.target as Element).id;
    if (id == 'D_' + ind) {
      if ((this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '') == '') {

        this.openDetails(Item, parentData, '1', '', tmp);
      } else {
        if (ref == '-') {
          Item.Dealer = '+';
        }
        if (ref == '+') {
          Item.Dealer = '-';
        }
      }
    }
    // //console.log(ind,id,Item.data2sign);
    if (id == 'DVN_' + ind) {
      if ((this.selectedDataGrouping.length == 3 ? this.selectedDataGrouping[2]?.columnname : '') == '') {
        this.openDetails(Item, parentData, '2', '', tmp);
      } else {
        if (ref == '-') {
          Item.data2sign = '+';
        }
        if (ref == '+') {
          Item.data2sign = '-';
        }
      }
    }
  }

  public getColorClass(value: number | null | undefined): string {
    if (value === null || value === undefined || value === 0) {
      return ''; // No class applied
    }
    return value > 0 ? 'positivebg' : 'negativebg';
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
  closePopup() {
    if (this.popupReference) {
      this.popupReference.close();
    }
  }
  Servicedetails: any = [];
  popupReference!: NgbModalRef;

  openDetails(Item: any, ParentItem: any, ref: any, subparent: any, temp: any) {

    if (ref == '1') {
      if (Item.data1 != undefined && Item.data1 != 'REPORTS TOTAL') {
        this.popupReference = this.shared.ngbmodal.open(
          temp,
          {
            // size:'xl',
            backdrop: 'static',
          }
        );
        this.Servicedetails = [
          {
            storeId: this.selectedDataGrouping[0]?.columnname == 'Store_Name' ? Item.StoreID : '',
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
            DepartmentS: this.Department.indexOf('Service') >= 0 ? 'S' : '',
            DepartmentP: this.Department.indexOf('Parts') >= 0 ? 'P' : '', // DepartmentP: '',
            layer: 1,
            db1value: this.selectedDataGrouping[0]?.columnname == 'Store_Name' ? Item.StoreID : Item.data1,
            details: this.Department.indexOf('Details') >= 0 ? 'D' : '',

            saletype: this.saletype.toString()

          },
        ];
      }
      if (Item.data1 != undefined && Item.data1 == 'REPORTS TOTAL') {
        this.popupReference = this.shared.ngbmodal.open(
          temp,
          {
            // size:'xl',
            backdrop: 'static',
          }
        );
        this.Servicedetails = [
          {
            storeId: '',
            // "SrvcName": Item.ServiceAdvisor_Name,
            StartDate: this.FromDate,
            EndDate: this.ToDate,
            var1: this.selectedDataGrouping[0]?.columnname,
            var2: '',
            var3: '',
            var1Value: this.selectedDataGrouping[0]?.columnname == 'Store_Name' ? this.storeIds : Item.data1,
            var2Value: '',
            var3Value: '',
            PaytypeC: this.Paytype[0] == 'Customerpay_0' ? 'C' : '',
            PaytypeW: this.Paytype[1] == 'Warranty_1' ? 'W' : '',
            PaytypeI: this.Paytype[2] == 'Internal_2' ? 'I' : '',
            DepartmentS: this.Department.indexOf('Service') >= 0 ? 'S' : '',
            DepartmentP: this.Department.indexOf('Parts') >= 0 ? 'P' : '', // DepartmentP: '',
            layer: 1,
            db1value: this.selectedDataGrouping[0]?.columnname == 'Store_Name' ? Item.StoreID : Item.data1,
            details: this.Department.indexOf('Details') >= 0 ? 'D' : '',

            saletype: this.saletype.toString()

          },
        ];
      }
    }
    // console.log(Item)
    if (ref == '2') {
      if (Item.data2 != undefined) {
        this.popupReference = this.shared.ngbmodal.open(
          temp,
          {
            // size:'xl',
            backdrop: 'static',
          }
        );
        this.Servicedetails = [
          {
            // "storeId": ParentItem.Store,
            //  "SrvcName": Item.ServiceAdvisor_Name,
            storeId: (this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '') == 'Store_Name' ? Item.StoreID : (this.selectedDataGrouping[0]?.columnname == 'Store_Name' && ParentItem.data1 != 'REPORTS TOTAL' ? ParentItem.StoreID : ''),
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
            DepartmentS: this.Department.indexOf('Service') >= 0 ? 'S' : '',
            DepartmentP: this.Department.indexOf('Parts') >= 0 ? 'P' : '', // DepartmentP: '',
            layer: 2,
            db1value: this.selectedDataGrouping[0]?.columnname == 'Store_Name' && ParentItem.data1 != 'REPORTS TOTAL' ? ParentItem.StoreID :
              (this.selectedDataGrouping[0]?.columnname == 'Store_Name' && ParentItem.data1 == 'REPORTS TOTAL' ? this.storeIds : ParentItem.data1),
            db2value: (this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '') == 'Store_Name' ? Item.StoreID : Item.data2,
            details: this.Department.indexOf('Details') >= 0 ? 'D' : '',

            saletype: this.saletype.toString()

          },
        ];
      }
    }
    if (ref == '3') {
      if (Item.data3 != undefined) {
        this.popupReference = this.shared.ngbmodal.open(
          temp,
          {
            // size:'xl',
            backdrop: 'static',
          }
        );
        this.Servicedetails = [
          {
            // "storeId": ParentItem.Store,
            //  "SrvcName": Item.ServiceAdvisor_Name,
            storeId: (this.selectedDataGrouping.length == 3 ? this.selectedDataGrouping[2]?.columnname : '') == 'Store_Name' ? Item.StoreID : (this.selectedDataGrouping[0]?.columnname == 'Store_Name' && Item.data1 != 'REPORTS TOTAL' ? Item.StoreID : ''),
            StartDate: this.FromDate,
            EndDate: this.ToDate,
            var1: this.selectedDataGrouping[0]?.columnname,
            var2: this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '',
            var3: (this.selectedDataGrouping.length == 3 ? this.selectedDataGrouping[2]?.columnname : ''),
            var1Value: Item.data1,
            var2Value: Item.data2,
            var3Value: Item.data3,
            PaytypeC: this.Paytype[0] == 'Customerpay_0' ? 'C' : '',
            PaytypeW: this.Paytype[1] == 'Warranty_1' ? 'W' : '',
            PaytypeI: this.Paytype[2] == 'Internal_2' ? 'I' : '',
            DepartmentS: Item.data2 == 'Service' ? 'S' : '',
            DepartmentP: Item.data2 == 'Parts' ? 'P' : '', // DepartmentP: '',
            layer: 3,
            db1value: this.selectedDataGrouping[0]?.columnname == 'Store_Name' ? ParentItem.StoreID : ParentItem.data1,
            db2value: (this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '') == 'Store_Name' ? subparent.StoreID : subparent.data2,
            db3value: (this.selectedDataGrouping.length == 3 ? this.selectedDataGrouping[2]?.columnname : '') == 'Store_Name' ? Item.StoreID : Item.data3,
            details: this.Department.indexOf('Details') >= 0 ? 'D' : '',

            saletype: this.saletype.toString()

          },
        ];
      }
    }
    this.ServicePersonDetails = []
    this.GetDetails()
  }

  viewRO(roData: any) {
    // const modalRef = this.ngbmodel.open(RepairOrderComponent, { size: 'md', windowClass: 'compModal' });
    // modalRef.componentInstance.data = { ro: roData.ASG_Ronumber, storeid: roData.storeid, vin: roData.vin, vehicleid: roData.vehicleid,custno: roData?.customernumber }; // Pass data to the modal component    
    // modalRef.result.then((result) => {
    //   console.log(result); // Handle modal close result
    // }, (reason) => {
    //   console.log(`Dismissed: ${reason}`); // Handle dismiss reason
    // });
  }
  ServicePersonDetails: any = [];
  details: any = [];
  NoDataDetails!: boolean;
  spinnerLoader: boolean = true;
  spinnerLoadersec: boolean = false;
  pageNumber: any = 0;
  GetDetails() {
    // this.shared.spinner.show()
    this.spinnerLoadersec = true
    const obj = {

      "startdate": this.Servicedetails[0].StartDate,
      "enddate": this.Servicedetails[0].EndDate,
      "StoreID": this.Servicedetails[0].storeId == undefined ? '' : this.Servicedetails[0].storeId,
      "PaytypeC": this.Servicedetails[0].PaytypeC,
      "PaytypeW": this.Servicedetails[0].PaytypeW,
      "PaytypeI": this.Servicedetails[0].PaytypeI,
      // "DepartmentS": this.Servicedetails[0].DepartmentS,
      // "DepartmentP": this.Servicedetails[0].DepartmentP,
      "DepartmentS": '',
      "DepartmentP": 'P',
      "RO_CloseDate": "",
      "var1": this.Servicedetails[0].var1 == 'Store_Name' ? 'storeid' : this.Servicedetails[0].var1,
      "var2": this.Servicedetails[0].var2 == 'Store_Name' ? 'storeid' : this.Servicedetails[0].var2,
      "var1Value": this.Servicedetails[0].db1value == undefined ? '' : this.Servicedetails[0].db1value,
      "var2Value": this.Servicedetails[0].db2value == undefined ? '' : this.Servicedetails[0].db2value,
      'var3': this.Servicedetails[0].var3 == 'Store_Name' ? 'storeid' : this.Servicedetails[0].var3,
      'var3Value': this.Servicedetails[0].db3value == undefined ? '' : this.Servicedetails[0].db3value,
      "type": "",
      "PageNumber": this.pageNumber,
      "PageSize": "100",
      "Details": this.Servicedetails[0].details,
      "saletype": this.Servicedetails[0].saletype
    };
    this.shared.api
      .postmethod(this.comm.routeEndpoint + 'GetServicesGrossSummaryDetailsGLV1', obj)
      .subscribe((res) => {
        this.spinnerLoader = false;
        this.spinnerLoadersec = false;
        if (res.status == 200) {
          this.details = res.response;
          this.ServicePersonDetails = [
            ...this.ServicePersonDetails,
            ...this.details,
          ];
          // //console.log(this.ServicePersonDetails);
          // this.shared.spinner.hide()

          if (this.ServicePersonDetails.length > 0) {
            this.NoDataDetails = false;
            this.spinnerLoadersec = false;
          } else {
            this.NoDataDetails = true;
            this.spinnerLoadersec = false;
          }
          this.scrollState = true
        }
      });
  }
  scrollState: boolean = false;
  Scrollpercent: any = 0;
  @ViewChild('scrollcent') scrollcent!: ElementRef;
  updateVerticalScroll(event: any): void {

    const scrollDemo = document.querySelector('#scrollcent') as HTMLElement;
    this.Scrollpercent = Math.round(
      (event.target.scrollTop / (event.target.scrollHeight - scrollDemo.clientHeight)) * 100);
    if (event.target.scrollTop + event.target.clientHeight >= event.target.scrollHeight - 2) {
      if (this.details.length == 100 && this.scrollState == true) {
        this.scrollState = false
        this.spinnerLoadersec = true;
        this.pageNumber++;
        this.GetDetails();
      }
    }
  }
  ngAfterViewInit() {

    this.shared.api.getStores().subscribe((res: any) => {
      if (this.comm.pageName == 'Parts Gross GL') {
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
        if (res.obj.title == 'Parts Gross GL') {
          if (res.obj.state == true) {
            this.exportToExcel();
          }
        }
      }
    });
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res) => {
      if (this.email != undefined) {
        if (res.obj.title == 'Parts Gross GL') {
          if (res.obj.stateEmailPdf == true) {
            // this.sendEmailData(res.obj.Email, res.obj.notes, res.obj.from);
          }
        }
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res) => {
      if (this.print != undefined) {
        if (res.obj.title == 'Parts Gross GL') {
          if (res.obj.statePrint == true) {
            // this.GetPrintData();
          }
        }
      }
    });
    this.Pdf = this.shared.api.getExportToPDFAllReports().subscribe((res) => {
      if (this.Pdf != undefined) {
        if (res.obj.title == 'Parts Gross GL') {
          if (res.obj.statePDF == true) {
            // this.generatePDF();
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
  multipleorsingle(block: any, e: any) {
    if (block == 'Dept') {
      const index = this.Department.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.Department.splice(index, 1);
        if (this.Department.length == 0) {
          this.toast.show('Please select atleast one Department', 'warning', 'Warning');
        }
      } else {
        this.Department.push(e);
      }
    }
    if (block == 'Parts') {
      if (e != 'All') {
        const index = this.saletype.findIndex((i: any) => i == e);
        if (index >= 0) {
          this.saletype.splice(index, 1);
          if (this.saletype.length == 0) {
            this.toast.show('Please select atleast one Saletype', 'warning', 'Warning');
          }
        } else {
          this.saletype.push(e);
        }
      }
      if (e == 'All') {
        if (this.saletype.length == this.partsSaleType.length) {
          this.saletype = []
        } else {
          this.saletype = this.partsSaleType.map(function (a: any) {
            return a.ASG_Subtype_Detail;
          });
        }
      }

    }
    if (block == 'TB') {
      this.TotalReport = e;

    }

  }
  partsSaleType: any = []
  getSaleTypeList() {
    const obj = {}
    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetPartsSaleTypes', obj).subscribe((res) => {
      if (res.response && res.response.length > 0) {
        console.log(res.response);
        this.partsSaleType = res.response.filter((e: any) => e.ASG_Subtype_Detail != '')
        if (this.saletype == '' || this.saletype == undefined) {
          this.saletype = this.partsSaleType.map(function (a: any) {
            return a.ASG_Subtype_Detail;
          });
        }
        // this.saletype = this.changes['header'].currentValue[0].saletype;

      }
    })
  }
  activePopover: number = -1;

  togglePopover(popoverIndex: number) {
    this.activePopover = this.activePopover === popoverIndex ? -1 : popoverIndex;
  }

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
  viewreport() {
    this.activePopover = -1
    if (this.selectedDataGrouping.length == 0) {
      this.toast.show('Please select atleast one Value from Grouping', 'warning', 'Warning');
    } else {
      if (this.storeIds.length == 0 && this.selectedotherstoreids.length == 0) {
        this.toast.show('Please select atleast one Store', 'warning', 'Warning');
      } else if (this.Department.length == 0) {
        this.toast.show('Please select atleast one Department Type', 'warning', 'Warning');
      }

      else {
        this.setHeaderData()
        this.getServiceData()
      }
    }
  }
  ExcelStoreNames: any = [];
  exportToExcel() {
    let storeNames: any = [];
    let storeIds = this.storeIds
    // const obj = {
    //   id: this.groups,
    //   userid: JSON.parse(localStorage.getItem('UserDetails')!).userid,
    // };
    // this.shared.api
    //   .postmethodOne(this.comm.routeEndpoint+'GetStoresbyGroupuserid', obj)
    //   .subscribe((res: any) => {
    storeNames = this.comm.groupsandstores
      .filter((v: any) => v.sg_id == this.groups)[0]
      .Stores.filter((item: any) =>
        storeIds.some((cat: any) => cat === item.ID.toString())
      );
    if (
      storeIds.length ==
      this.comm.groupsandstores.filter((v: any) => v.sg_id == this.groups)[0]
        .Stores.length
    ) {
      this.ExcelStoreNames = 'All Stores';
    } else {
      this.ExcelStoreNames = storeNames.map(function (a: any) {
        return a.storename;
      });
    }
    const ServiceData = this.ServiceData.map((_arrayElement: any) =>
      Object.assign({}, _arrayElement)
    );
    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('Parts Gross GL');
    worksheet.views = [
      {
        showGridLines: false,
      },
    ];
    worksheet.addRow('');
    const titleRow = worksheet.addRow(['Parts Gross GL']);
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
    const lastyearDate = this.shared.datePipe.transform(this.lastyearDate, 'MMM YYYY');
    const DateToday = this.shared.datePipe.transform(
      new Date(),
      'MM/dd/yyyy h:mm:ss a'
    );
    const DATE_EXTENSION = this.shared.datePipe.transform(new Date(), 'MMddyyyy');
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
    groups.value = this.comm.groupsandstores.filter(
      (val: any) => val.sg_id == this.groups.toString()
    )[0].sg_name;
    groups.font = { name: 'Arial', family: 4, size: 9 };
    const Stores1 = worksheet.getCell('B10');
    Stores1.value = 'Stores :';
    const stores1 = worksheet.getCell('D10');
    stores1.value =
      this.ExcelStoreNames == 0
        ? 'All Stores'
        : this.ExcelStoreNames == null
          ? '-'
          : this.ExcelStoreNames.toString().replaceAll(',', ', ');
    stores1.font = { name: 'Arial', family: 4, size: 9 };
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

    let saleblock = worksheet.getCell('B19');
    saleblock.value = 'Sale';
    saleblock.alignment = { vertical: 'middle', horizontal: 'center' };
    saleblock.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
      color: { argb: 'FFFFFF' },
    };
    saleblock.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '2a91f0' },
      bgColor: { argb: 'FF0000FF' },
    };
    saleblock.border = { right: { style: 'dotted' } };

    worksheet.mergeCells('C19', 'I19');
    let units = worksheet.getCell('C19');
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
    worksheet.mergeCells('J19', 'M19');
    let frontgross = worksheet.getCell('J19');
    frontgross.value = 'Volume';
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
    worksheet.mergeCells('N19', 'P19');
    let backgross = worksheet.getCell('N19');
    backgross.value = 'Performance';
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
    backgross.border = { right: { style: 'thin' } };
    let Headings = [
      FromDate + ' - ' + ToDate + ', ' + PresentYear,
      'Total',
      'Total',
      'Pace',
      'Target',
      'Diff',
      lastyearDate,
      'Diff',
      'YOY%',
      'Total',
      'Pace',
      'Target',
      'Diff',
      'Sales/RO',
      'Avg RO',
      'GP%',
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
      var obj = [
        d.data1 == '' ? '-' : d.data1 == null ? '-' : d.data1,
        d.totalsale == '' ? '-' : d.totalsale == null ? '-' : d.totalsale,
        d.Total == '' ? '-' : d.Total == null ? '-' : d.Total,
        d.TotalGross_Pace == '' ? '-' : d.TotalGross_Pace == null ? '-' : d.TotalGross_Pace,
        d.Gross_Target == '' ? '-' : d.Gross_Target == null ? '-' : d.Gross_Target,
        d.Diff == '' ? '-' : d.Diff == null ? '-' : d.Diff,
        d.Lmonth == '' ? '-' : d.Lmonth == null ? '-' : d.Lmonth,
        d.LDiff == '' ? '-' : d.LDiff == null ? '-' : d.LDiff,
        d.YOY == '' ? '-' : d.YOY == null ? '-' : d.YOY + '%',
        d.RoToal == '' ? '-' : d.RoToal == null ? '-' : d.RoToal,
        d.Total_ROPACE == '' ? '-' : d.Total_ROPACE == null ? '-' : d.Total_ROPACE,
        d.RO_Target == '' ? '-' : d.RO_Target == null ? '-' : d.RO_Target,
        d.ROCountDiff == '' ? '-' : d.ROCountDiff == null ? '-' : d.ROCountDiff,
        d.Sales_Per_RO == '' ? '-' : d.Sales_Per_RO == null ? '-' : d.Sales_Per_RO,
        d.Average_RO == '' ? '-' : d.Average_RO == null ? '-' : d.Average_RO,
        d.GP == '' ? '-' : d.GP == null ? '-' : d.GP + ' %',
      ];
      const Data1 = worksheet.addRow(obj);
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
          (number > 1 && number < 9) || number == 15 || number == 14) {
          cell.numFmt = '$#,##0';
        }

        if (number == 9 || number == 10 || number == 11 || number == 12) {
          cell.numFmt = '#,##0';
        }
        if (number != 1) {
          cell.alignment = { vertical: 'top', horizontal: 'center', indent: 1 };
        }
        if (obj[number] < 0) {
          Data1.getCell(number + 1).font = {
            name: 'Arial',
            family: 4,
            size: 9,
            color: { argb: 'FFFF0000' },
          };
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
          let obj2 = [
            d1.data2 == '' ? '-' : d1.data2 == null ? '-' : d1.data2,
            d1.totalsale == '' ? '-' : d1.totalsale == null ? '-' : d1.totalsale,

            d1.Total == '' ? '-' : d1.Total == null ? '-' : d1.Total,
            d1.TotalGross_Pace == '' ? '-' : d1.TotalGross_Pace == null ? '-' : d1.TotalGross_Pace,
            d1.Gross_Target == '' ? '-' : d1.Gross_Target == null ? '-' : d1.Gross_Target,
            d1.Diff == '' ? '-' : d1.Diff == null ? '-' : d1.Diff,

            d1.Lmonth == '' ? '-' : d1.Lmonth == null ? '-' : d1.Lmonth,
            d1.LDiff == '' ? '-' : d1.LDiff == null ? '-' : d1.LDiff,
            d1.YOY == '' ? '-' : d1.YOY == null ? '-' : d1.YOY + '%',
            d1.RoToal == '' ? '-' : d1.RoToal == null ? '-' : d1.RoToal,
            d1.Total_ROPACE == '' ? '-' : d1.Total_ROPACE == null ? '-' : d1.Total_ROPACE,
            '-',
            '-',
            d1.Sales_Per_RO == '' ? '-' : d1.Sales_Per_RO == null ? '-' : d1.Sales_Per_RO,
            d1.Average_RO == '' ? '-' : d1.Average_RO == null ? '-' : d1.Average_RO,
            d1.GP == '' ? '-' : d1.GP == null ? '-' : d1.GP + ' %',
          ];
          const Data2 = worksheet.addRow(obj2);
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
              (number > 1 && number < 9) || number == 15 || number == 14) {
              cell.numFmt = '$#,##0.00';
            }

            if (number == 9 || number == 10 || number == 11 || number == 12) {
              cell.numFmt = '#,##0';
            }
            if (number != 1) {
              cell.alignment = { vertical: 'top', horizontal: 'center', indent: 1 };
            }
            if (number != 1) {
              cell.alignment = {
                vertical: 'top',
                horizontal: 'center',
                indent: 1,
              };
            }
            if (obj2[number] < 0) {
              Data2.getCell(number + 1).font = {
                name: 'Arial',
                family: 4,
                size: 9,
                color: { argb: 'FFFF0000' },
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
          if (d1.subdata != undefined) {
            for (const d2 of d1.subdata) {
              let obj3 = [
                d2.data3 == '' ? '-' : d2.data3 == null ? '-' : d2.data3,
                d2.totalsale == '' ? '-' : d2.totalsale == null ? '-' : d2.totalsale,
                d2.Total == '' ? '-' : d2.Total == null ? '-' : d2.Total,
                d2.TotalGross_Pace == '' ? '-' : d2.TotalGross_Pace == null ? '-' : d2.TotalGross_Pace,
                d2.Gross_Target == '' ? '-' : d2.Gross_Target == null ? '-' : d2.Gross_Target,
                d2.Diff == '' ? '-' : d2.Diff == null ? '-' : d2.Diff,
                d2.Lmonth == '' ? '-' : d2.Lmonth == null ? '-' : d2.Lmonth,
                d2.LDiff == '' ? '-' : d2.LDiff == null ? '-' : d2.LDiff,
                d2.YOY == '' ? '-' : d2.YOY == null ? '-' : d2.YOY + '%',
                d2.RoToal == '' ? '-' : d2.RoToal == null ? '-' : d2.RoToal,
                d2.Total_ROPACE == '' ? '-' : d2.Total_ROPACE == null ? '-' : d2.Total_ROPACE,
                '-',
                '-',
                d2.Sales_Per_RO == '' ? '-' : d2.Sales_Per_RO == null ? '-' : d2.Sales_Per_RO,
                d2.Average_RO == '' ? '-' : d2.Average_RO == null ? '-' : d2.Average_RO,
                d2.GP == '' ? '-' : d2.GP == null ? '-' : d2.GP + ' %',
              ];
              const Data3 = worksheet.addRow(obj3);
              Data3.outlineLevel = 2; // Grouping level 2
              Data3.font = { name: 'Arial', family: 4, size: 8 };
              Data3.alignment = { vertical: 'middle', horizontal: 'center' };
              Data3.getCell(1).alignment = {
                indent: 3,
                vertical: 'middle',
                horizontal: 'left',
              };
              Data3.eachCell((cell: any, number: any) => {
                cell.border = { right: { style: 'dotted' } };
                cell.numFmt = '$#,##0';
                if (
                  (number > 1 && number < 9) || number == 15 || number == 14) {
                  cell.numFmt = '$#,##0';
                }

                if (number == 9 || number == 10 || number == 11 || number == 12) {
                  cell.numFmt = '#,##0';
                }
                if (number != 1) {
                  cell.alignment = { vertical: 'top', horizontal: 'center', indent: 1 };
                }
                if (number != 1) {
                  cell.alignment = {
                    vertical: 'top',
                    horizontal: 'center',
                    indent: 1,
                  };
                }
                if (obj3[number] < 0) {
                  Data3.getCell(number + 1).font = {
                    name: 'Arial',
                    family: 4,
                    size: 9,
                    color: { argb: 'FFFF0000' },
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
        if (rowIndex > 1 && rowIndex < 19) {
          // Skip the header row
          // Apply conditional alignment based on your conditions
          if (colIndex === 1) {
            // Apply right alignment to the second column
            cell.alignment = {
              horizontal: 'left',
              vertical: 'middle',
              indent: 1,
            };
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
    worksheet.getColumn(8).width = 25;
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
    workbook.xlsx.writeBuffer().then((data: any) => {
      const blob = new Blob([data], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      this.shared.exportToExcel(workbook, 'Parts Gross GL')
    });
    // });
  }

  exportAsXLSX() {
    const obj = {
      "startdate": this.Servicedetails[0].StartDate,
      "enddate": this.Servicedetails[0].EndDate,
      "StoreID": this.Servicedetails[0].storeId == undefined ? '' : this.Servicedetails[0].storeId,
      "PaytypeC": this.Servicedetails[0].PaytypeC,
      "PaytypeW": this.Servicedetails[0].PaytypeW,
      "PaytypeI": this.Servicedetails[0].PaytypeI,
      // "DepartmentS": this.Servicedetails[0].DepartmentS,
      // "DepartmentP": this.Servicedetails[0].DepartmentP,
      "DepartmentS": this.Servicedetails[0].DepartmentS,
      "DepartmentP": this.Servicedetails[0].DepartmentP,

      "RO_CloseDate": "",
      "var1": this.Servicedetails[0].var1 == 'Store_Name' ? 'storeid' : this.Servicedetails[0].var1,
      "var2": this.Servicedetails[0].var2 == 'Store_Name' ? 'storeid' : this.Servicedetails[0].var2,
      "var1Value": this.Servicedetails[0].db1value == undefined ? '' : this.Servicedetails[0].db1value,
      "var2Value": this.Servicedetails[0].db2value == undefined ? '' : this.Servicedetails[0].db2value,
      'var3': this.Servicedetails[0].var3 == 'Store_Name' ? 'storeid' : this.Servicedetails[0].var3,
      'var3Value': this.Servicedetails[0].db3value == undefined ? '' : this.Servicedetails[0].db3value,
      "type": "",
      "PageNumber": 0,
      "PageSize": "1000000",
      "Details": this.Servicedetails[0].details,
      "saletype": this.Servicedetails[0].saletype
    };
    this.shared.spinner.show()
    this.shared.api
      .postmethod(this.comm.routeEndpoint + 'GetServicesGrossSummaryDetailsGLV1', obj)
      .subscribe((res) => {
        if (res.status == 200) {
          this.shared.spinner.hide()
          let localarray = res.response
          const workbook = this.shared.getWorkbook();
          const worksheet = workbook.addWorksheet(this.Servicedetails[0].screen + 'Details');
          worksheet.views = [
            {
              state: 'frozen',
              ySplit: 8, // Number of rows to freeze (2 means the first two rows are frozen)
              topLeftCell: 'A10', // Specify the cell to start freezing from (in this case, the third row)
              showGridLines: false,
            },
          ];
          const DateToday = this.shared.datePipe.transform(new Date(), 'MM/dd/yyyy h:mm:ss a');
          const titleRow = worksheet.getCell("A2"); titleRow.value = this.Servicedetails[0].screen + 'Details';
          titleRow.font = { name: 'Arial', family: 4, size: 15, bold: true };
          titleRow.alignment = { indent: 1, vertical: 'middle', horizontal: 'left' }
          const DateBlock = worksheet.getCell("L2"); DateBlock.value = DateToday;
          DateBlock.font = { name: 'Arial', family: 4, size: 10 };
          DateBlock.alignment = { vertical: 'middle', horizontal: 'center' }
          worksheet.addRow([''])
          const Store_Name = worksheet.addRow(['Store Name :']);
          Store_Name.getCell(1).font = { name: 'Arial', family: 4, size: 9, bold: true, };
          Store_Name.alignment = { indent: 1, vertical: 'middle', horizontal: 'left' }
          const StoreName = worksheet.getCell("B4"); StoreName.value = this.Servicedetails[0].var1Value;
          StoreName.font = { name: 'Arial', family: 4, size: 9 };
          StoreName.alignment = { indent: 1, vertical: 'middle', horizontal: 'left' }
          const DATE_EXTENSION = this.shared.datePipe.transform(
            new Date(),
            'MMddyyyy'
          );
          const StartDealDate = worksheet.addRow(['Start Date :']);
          const startdealdate = worksheet.getCell('B5');
          startdealdate.value = this.Servicedetails[0].StartDate;
          startdealdate.font = { name: 'Arial', family: 4, size: 9 };
          startdealdate.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
          StartDealDate.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
          StartDealDate.getCell(1).font = {
            name: 'Arial',
            family: 4,
            size: 9,
            bold: true,
          };
          const EndDealDate = worksheet.addRow(['End Date :']);
          EndDealDate.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
          const enddealdate = worksheet.getCell('B6');
          enddealdate.value = this.Servicedetails[0].EndDate;
          enddealdate.font = { name: 'Arial', family: 4, size: 9 };
          enddealdate.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
          EndDealDate.getCell(1).font = {
            name: 'Arial',
            family: 4,
            size: 9,
            bold: true,
          };

          worksheet.addRow('');
          let Headings = [
            'Sl.no',
            'RO #',
            'Department',
            'Account #',
            'Account Type',
            'Account Type Detail',
            'Description',
            'Sale Type',
            'Posting Amount',
            'Accounting Date'
          ];
          const headerRow = worksheet.addRow(Headings);
          headerRow.font = {
            name: 'Arial',
            family: 4,
            size: 9,
            bold: true,
            color: { argb: 'FFFFFF' },
          };
          headerRow.height = 20;
          headerRow.alignment = {
            indent: 1,
            vertical: 'middle',
            horizontal: 'center',
          };
          headerRow.eachCell((cell, number) => {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: '2a91f0' },
              bgColor: { argb: 'FF0000FF' },
            };
            headerRow.height = 20;
            cell.border = { right: { style: 'thin' } };
            cell.alignment = {
              vertical: 'middle',
              horizontal: 'center',
              wrapText: true,
            };
          });
          var count = 0
          for (const d of localarray) {
            count++
            d.ASG_accountingdate = this.shared.datePipe.transform(d.ASG_accountingdate, 'MM/dd/yyyy');
            let obj = [count,
              d.ASG_Ronumber == '' ? '-' : d.ASG_Ronumber == null ? '-' : d.ASG_Ronumber,
              d.ASG_Department == '' ? '-' : d.ASG_Department == null ? '-' : d.ASG_Department,
              d.ASG_AccountNumber == '' ? '-' : d.ASG_AccountNumber == null ? '-' : d.ASG_AccountNumber,
              d.ASG_AccountType == '' ? '-' : d.ASG_AccountType == null ? '-' : d.ASG_AccountType,
              d.ASG_Account_type_detail == '' ? '-' : d.ASG_Account_type_detail == null ? '-' : d.ASG_Account_type_detail,
              d.ASG_AccountFullDetail == '' ? '-' : d.ASG_AccountFullDetail == null ? '-' : d.ASG_AccountFullDetail,
              d.ASG_Subtype_Detail == '' ? '-' : d.ASG_Subtype_Detail == null ? '-' : d.ASG_Subtype_Detail,
              d.ASG_postingamount == '' ? '-' : d.ASG_postingamount == null ? '-' : d.ASG_postingamount,
              d.ASG_accountingdate == '' ? '-' : d.ASG_accountingdate == null ? '-' : d.ASG_accountingdate,]
            const Data1 = worksheet.addRow(obj);
            // Data1.outlineLevel = 1; // Grouping level 1
            Data1.font = { name: 'Arial', family: 4, size: 8 };
            Data1.height = 18;
            // Data1.getCell(1).alignment = {indent: 1,vertical: 'middle', horizontal: 'left'}
            Data1.eachCell((cell, number) => {
              cell.border = { right: { style: 'thin' } };
              // cell.numFmt = '$#,##0';
              cell.alignment = { indent: 1, vertical: 'middle', horizontal: 'center', };

              if (number == 9) {
                cell.numFmt = '$#,##0';
                cell.alignment = { indent: 1, vertical: 'middle', horizontal: 'center', };
              }
              if (obj[number] < 0) {
                Data1.getCell(number + 1).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } };

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
          }
          worksheet.getColumn(1).width = 18;
          worksheet.getColumn(2).width = 20;
          worksheet.getColumn(3).width = 10;
          worksheet.getColumn(4).width = 15;
          worksheet.getColumn(5).width = 20;
          worksheet.getColumn(6).width = 20;
          worksheet.getColumn(7).width = 30;
          worksheet.getColumn(8).width = 15;
          worksheet.getColumn(9).width = 20;
          worksheet.getColumn(10).width = 15;

          worksheet.addRow([]);
          workbook.xlsx.writeBuffer().then((data: any) => {
            const blob = new Blob([data], {
              type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            });
            this.shared.exportToExcel(workbook, 'Parts Gross GL Details')
          });
        }
      });
  }

}