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
import { Options, LabelType } from "@angular-slider/ngx-slider";
import { NgxSliderModule } from '@angular-slider/ngx-slider';
import { ServiceOpenRODetails } from '../service-open-rodetails/service-open-rodetails';
@Component({
  selector: 'app-dashboard',
  imports: [SharedModule, BsDatepickerModule, DateRangePicker, Stores,ServiceOpenRODetails,NgxSliderModule],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
})
export class Dashboard {
  scaleValue: any = 1;
  spinnerLoadersec: boolean = false;
  ServiceData: any = []
  IndividualServiceGross: any = [];
  TotalServiceGross: any = [];
  type: any = 'A';

  TotalReport: string = 'T';
  NoData: boolean = false;
  ROType: any = 'Closed';
  Department: any = ['Service', 'Parts'];
  Paytype: any = ['Customerpay_0', 'Warranty_1', 'Internal_2', 'Extended_3'];
  Grosstype: any = ['Labour_0', '', 'Misc_2', 'Sublet_3'];

  responcestatus: any = '';
  CurrentDate = new Date();
  groups: any = 1;
  AgeFrom: any = 1;
  AgeTo: any = 1000;
  inventory: any = 'All'
  rostatus: any = ['All'];
  topfive: boolean = false;
  otherstoreid: any = '';
  selectedotherstoreids: any = '';

  reportOpenSub!: Subscription;
  reportGetting!: Subscription;
  Pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;


  asofnow: any = '';
  loadbtnFlag = ''
  selectedDataGrouping: any = [];
  dataGrouping: any = [
    { "ARG_ID": 40, "ARG_LABEL": "Store", "ARG_SEQ": 0, "state": true, "columnname": "Store_Name", "Active": "Y" },
    { "ARG_ID": 26, "ARG_LABEL": "Advisor Name", "ARG_SEQ": 1, "state": false, "Active": "Y", "columnname": "ServiceAdvisor_Name" },
    { "ARG_ID": 27, "ARG_LABEL": "Advisor Number", "ARG_SEQ": 2, "state": false, "Active": "Y", "columnname": "ServiceAdvisor" },
    { "ARG_ID": 28, "ARG_LABEL": "Tech Name", "ARG_SEQ": 3, "state": false, "Active": "N", "columnname": "" },
    { "ARG_ID": 29, "ARG_LABEL": "Tech Number", "ARG_SEQ": 4, "state": false, "Active": "N", "columnname": "" },
    { "ARG_ID": 30, "ARG_LABEL": "Pay Type", "ARG_SEQ": 5, "state": false, "Active": "N", "columnname": "" },
    { "ARG_ID": 31, "ARG_LABEL": "Vehicle Year", "ARG_SEQ": 6, "state": false, "Active": "Y", "columnname": "vehicle_Year" },
    { "ARG_ID": 32, "ARG_LABEL": "Vehicle Make", "ARG_SEQ": 7, "state": false, "Active": "Y", "columnname": "vehicle_Make" },
    { "ARG_ID": 33, "ARG_LABEL": "Vehicle Model", "ARG_SEQ": 8, "state": false, "Active": "Y", "columnname": "Vehicle_Model" },
    { "ARG_ID": 34, "ARG_LABEL": "Vehicle Odometer", "ARG_SEQ": 9, "state": false, "Active": "Y", "columnname": "Vehicle_Odometer" },
    { "ARG_ID": 35, "ARG_LABEL": "Customer Name", "ARG_SEQ": 10, "state": false, "Active": "Y", "columnname": "CName" },
    { "ARG_ID": 36, "ARG_LABEL": "Customer ZIP", "ARG_SEQ": 11, "state": false, "Active": "Y", "columnname": "CZip" },
    { "ARG_ID": 37, "ARG_LABEL": "Customer State", "ARG_SEQ": 12, "state": false, "Active": "Y", "columnname": "CState" },
    { "ARG_ID": 38, "ARG_LABEL": "RO Open Date", "ARG_SEQ": 13, "state": false, "Active": "Y", "columnname": "opendate" },
    { "ARG_ID": 39, "ARG_LABEL": "No Grouping", "ARG_SEQ": 14, "state": false, "Active": "N", "columnname": "" },
    { "ARG_ID": 68, "ARG_LABEL": "Opcode", "ARG_SEQ": 15, "state": false, "Active": "Y" }
  ]

    optionProx: Options = {
    floor: 0, ceil: 1000,
    showSelectionBarFromValue: 0,
    hideLimitLabels: true,
    translate: (value: number, label: LabelType): string => {
      switch (label) {
        case LabelType.Low: return "<a>" + value + " Age</a>";
        case LabelType.High:
          return " " + value + " Age";
        default:
          return "" + value + " Age";
      }
    }
  };

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
    'FromDate': this.FromDate, 'ToDate': this.ToDate, "MaxDate": this.maxDate, 'MinDate': this.minDate, 'DateType': this.DateType, 'DisplayTime': this.displaytime,custom:'N',
    Types: [
      { 'code': 'Overall', 'name': "All Open RO's" },
      { 'code': '3', 'name': '> 3 Days' },
      { 'code': '10', 'name': '> 10 Days' },

    ]
  }

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
    this.initializeDates('Overall')
    if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
      this.groupId = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Preferences
      this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids.split(',')
    }
    if (this.shared.common.groupsandstores.length > 0) {
      this.groupsArray = this.shared.common.groupsandstores.filter((val: any) => val.sg_id != this.shared.common.reconID);
      this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
      this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_Name : this.groupName = ''
      this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
      // this.comm.redirectionFrom.flag == 'V' ? this.rostatus = this.comm.redirectionFrom.ro_filter : this.rostatus = 'All';

      this.getStoresandGroupsValues()
    }
    this.shared.setTitle(this.comm.titleName + '-Service Open RO');


    this.setHeaderData()
    this.getServiceData();

  }
  ngOnInit(): void {
    this.RunServiceOpenLoad('')

  }

  setHeaderData() {
    const data = {
      title: 'Service Open RO',
      stores: this.storeIds,
      ROType: this.ROType,
      Department: this.Department,
      Paytype: this.Paytype,

      ToporBottom: this.TotalReport,
      fromdate: this.FromDate,
      todate: this.ToDate,
      grosstype: this.Grosstype,

      groups: this.groupId,
      AgeFrom: this.AgeFrom,
      AgeTo: this.AgeTo,
      inventory: this.inventory,
      rostatus: this.rostatus,

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
  Load(action: any) {
    this.shared.spinner.show()
    this.RunServiceOpenLoad(action)
  }
  RunServiceOpenLoad(action: any) {
    const obj = {
      "ACTION": action
    }
    this.shared.api.postmethod(this.comm.routeEndpoint + 'RunServiceOpenLoad', obj).subscribe((res: any) => {
      if (res && res.response) {
        console.log(res.response);
        if (action == '') {
          this.asofnow = res.response.substring(2);
          this.loadbtnFlag = res.response.substring(0, 1);
        } else {
          this.RunServiceOpenLoad('')
          this.getServiceData()
        }

      }
    })
  }
  getServiceData() {
    if (this.storeIds != '' || this.selectedotherstoreids != '') {
      this.responcestatus = '';
      this.shared.spinner.show();
      this.GetData();
      this.GetTotalData();
    } else {
      // this.NoData = true
    }
  }
  GetData() {
    this.IndividualServiceGross = [];
    const obj = {
      startdate: this.FromDate,
      enddate: this.ToDate,
      // enddate: "11-20-2023",
      StoreID: this.selectedotherstoreids != undefined && this.selectedotherstoreids != '' && this.selectedotherstoreids != null ?
        (this.storeIds != '' ? this.storeIds.toString() + ',' + this.selectedotherstoreids.toString() : this.selectedotherstoreids.toString()) : this.storeIds.toString(),
      AdvisorNumber: '',
      AdvisorName: '',
      ROSTATUS: this.rostatus == 'All' ? '' : this.rostatus.toString(),
      PaytypeCP: this.Paytype[0] == 'Customerpay_0' ? 'Y' : '',
      PaytypeWarranty: this.Paytype[1] == 'Warranty_1' ? 'Y' : '',
      PaytypeInternal: this.Paytype[2] == 'Internal_2' ? 'Y' : '',
      PaytypeExtendedWarranty: this.Paytype[3] == 'Extended_3' ? 'Y' : '',
      Department: '',
      GrossTypeLabor: this.Department.indexOf('Service') >= 0 ? 'Y' : '',
      GrossTypeParts: this.Department.indexOf('Parts') >= 0 ? 'Y' : '',
      GrossTypeMisc: '',
      GrossTypeSublet: '',
      // GrossTypeMisc: this.Grosstype[2] == 'Misc_2' ? 'Y' : '',
      // GrossTypeSublet: this.Grosstype[3] == 'Sublet_3' ? 'Y' : '',
      vehicle_Year: '',
      vehicle_Make: '',
      Vehicle_Model: '',
      Vehicle_Odometer: '',
      CName: '',
      CZip: '',
      CState: '',
      RO_OpenDate: '',
      Inventory: this.inventory == 'All' ? '' : this.inventory,
      var1: this.selectedDataGrouping.length >= 1 ? this.selectedDataGrouping[0]?.columnname : '',
      var2: this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '',
      var3: this.selectedDataGrouping.length == 3 ? this.selectedDataGrouping[2]?.columnname : '',
      type: 'D',
      minage: this.AgeFrom,
      maxage: this.AgeTo,
      Oldro: this.topfive == true ? 'Y' : ''
    };
    const curl = environment.apiUrl + this.comm.routeEndpoint + 'GetServiceSummaryBetaOpen';
    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetServiceSummaryBetaOpen', obj).subscribe(
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
              let topfive = this.topfive;
              let length = this.IndividualServiceGross.length
              this.IndividualServiceGross.some(function (x: any) {
                if (x.RoInfo != undefined) {
                  x.RoInfo = JSON.parse(x.RoInfo);
                  x.Dealer = '+'
                } else {
                  x.Dealer = '+'
                }
                // if (x.Comments != undefined && x.Comments != null) {
                //   x.Comments = JSON.parse(x.Comments);
                // }
                if (x.Data2 != undefined) {
                  x.Data2 = JSON.parse(x.Data2);
                  x.Data2 = x.Data2.map((v: any) => ({
                    ...v,
                    SubData: [],
                    data2sign: '+',
                  }));
                }
                if (path2 == '' && length == 1) {
                  x.Dealer = '-';
                }
                else if (path2 != '') {
                  x.Dealer = '-'
                }
              });
              this.combineIndividualandTotal();

            } else {
              // this.toast.error('Empty Response','');
              this.shared.spinner.hide();
              this.NoData = true;
            }
          } else {
            // this.toast.error('Empty Response','');
            this.shared.spinner.hide();
            this.NoData = true;
          }
        } else {
          //this.toast.error('No Response', '');
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
      // enddate: "11-20-2023",
      StoreID: this.selectedotherstoreids != undefined && this.selectedotherstoreids != '' && this.selectedotherstoreids != null ?
        (this.storeIds != '' ? this.storeIds.toString() + ',' + this.selectedotherstoreids.toString() : this.selectedotherstoreids.toString()) : this.storeIds.toString(),
      AdvisorNumber: '',
      AdvisorName: '',
      ROSTATUS: this.rostatus == 'All' ? '' : this.rostatus.toString(),
      PaytypeCP: this.Paytype[0] == 'Customerpay_0' ? 'Y' : '',
      PaytypeWarranty: this.Paytype[1] == 'Warranty_1' ? 'Y' : '',
      PaytypeInternal: this.Paytype[2] == 'Internal_2' ? 'Y' : '',
      PaytypeExtendedWarranty: this.Paytype[3] == 'Extended_3' ? 'Y' : '',

      Department: '',
      // GrossTypeLabor: this.Grosstype[0] == 'Labour_0' ? 'Y' : '',
      // GrossTypeParts: this.Grosstype[1] == 'Parts_1' ? 'Y' : '',
      // GrossTypeMisc: this.Grosstype[2] == 'Misc_2' ? 'Y' : '',
      // GrossTypeSublet: this.Grosstype[3] == 'Sublet_3' ? 'Y' : '',
      GrossTypeLabor: this.Department.indexOf('Service') >= 0 ? 'Y' : '',
      GrossTypeParts: this.Department.indexOf('Parts') >= 0 ? 'Y' : '',
      GrossTypeMisc: '',
      GrossTypeSublet: '',
      vehicle_Year: '',
      vehicle_Make: '',
      Vehicle_Model: '',
      Vehicle_Odometer: '',
      CName: '',
      CZip: '',
      CState: '',
      RO_OpenDate: '',
      Inventory: this.inventory == 'All' ? '' : this.inventory,
      var1: this.selectedDataGrouping[0]?.columnname == 'Store_Name' ? '' : this.selectedDataGrouping[0]?.columnname,
      var2: this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '',
      var3: this.selectedDataGrouping.length == 3 ? this.selectedDataGrouping[2]?.columnname : '',
      type: 'T',
      minage: this.AgeFrom,
      maxage: this.AgeTo,
      Oldro: this.topfive == true ? 'Y' : ''
    };
    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetServiceSummaryBetaOpen', obj).subscribe(
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
              this.combineIndividualandTotal();
            } else {
              // this.toast.error('Empty Response','');
              this.shared.spinner.hide();
              this.NoData = true;
            }
          } else {
            // this.toast.error('Empty Response','');
            this.shared.spinner.hide();
            this.NoData = true;
          }
        } else {
          //this.toast.error('No Response', '');
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
        if (this.selectedDataGrouping.length == 1 && this.ServiceData.length == 2) {
          console.log(this.ServiceData, this.topfive);
          setTimeout(() => {
            // if(this.topfive == false)
            (<HTMLInputElement>document.getElementById('D_1')).click();
          }, 300);
        }
        if (this.topfive == true) {
          setTimeout(() => {
            this.ServiceData.forEach((val: any, i: any) => {
              if (i < this.ServiceData.length - 1) {
                setTimeout(() => {
                  val.Dealer = '-'
                  if (this.selectedDataGrouping.length == 1) {
                    this.openRODetails(val);
                  } else {
                    val.Data2.forEach((sub: any, j: any) => {
                      sub.data2sign = '-'
                      this.openRODetails(sub, val);
                    })
                  }
                }, 100);
              }
            })
          }, 300);
        }
      } else {
        this.IndividualServiceGross.unshift(this.TotalServiceGross[0]);
        this.ServiceData = this.IndividualServiceGross;
        if (this.selectedDataGrouping.length == 1 && this.ServiceData.length == 2) {
          setTimeout(() => {
            // if(this.topfive == false)
            (<HTMLInputElement>document.getElementById('D_1')).click();
          }, 300);
          // alert('Hi')
        }
        if (this.topfive == true) {
          this.ServiceData.forEach((val: any, i: any) => {
            if (i > 0) {
              setTimeout(() => {
                val.Dealer = '-'
                if (this.selectedDataGrouping.length == 1) {
                  this.openRODetails(val);
                } else {
                  val.Data2.forEach((sub: any, j: any) => {
                    sub.data2sign = '-'
                    this.openRODetails(sub, val);
                  })
                }
              }, 100);
            }
          })
        }
      }
      // if (this.path2 == '' && this.IndividualServiceGross.length == 1) {
      //   // this.IndividualServiceGross[0].Dealer = '-';
      //   setTimeout(() => {
      //   }, 200);
      //   //
      // } else {
      //   // x.Dealer = '+';
      // }
      this.shared.spinner.hide();
      console.log(this.ServiceData);
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
  toggleAction: any = '-';
  expandorcollapse(ind: any, e: any, ref: any, Item: any, parentData: any) {
    let id = (e.target as Element).id;
    this.toggleAction = ref;
    if (id == 'D_' + ind) {
      if (ref == '-') {
        Item.Dealer = '+';
      }
      if (ref == '+') {
        Item.Dealer = '-';
      }
      if (this.selectedDataGrouping.length == 1) {
        this.openRODetails(Item, Item, '1', '');
      }
    }
    if (id == 'DN_' + ind) {
      if (ref == '-') {
        Item.data2sign = '+';
      }
      if (ref == '+') {
        Item.data2sign = '-';
        Item.Dealer = '-';
        this.openRODetails(Item, parentData, '1', '');
      }
    }
  }
  openRODetails(Item?: any, ParentItem?: any, subParentItem?: any, ref?: any) {
    if (this.topfive == true) {
      this.RODetailsObjectMain = [
        {
          StartDate: this.FromDate,
          EndDate: this.ToDate,
          var1: this.selectedDataGrouping[0]?.columnname,
          var2: this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '',
          GrossTypeLabor: this.Department.indexOf('Service') >= 0 ? 'Y' : '',
          GrossTypeParts: this.Department.indexOf('Parts') >= 0 ? 'Y' : '',
          GrossTypeMisc: this.Department.indexOf('Misc') >= 0 ? 'Y' : '',
          GrossTypeSublet: this.Department.indexOf('Sublet') >= 0 ? 'Y' : '',
          PaytypeCP: this.Paytype[0] == 'Customerpay_0' ? 'Y' : '',
          PaytypeWarranty: this.Paytype[1] == 'Warranty_1' ? 'Y' : '',
          PaytypeInternal: this.Paytype[2] == 'Internal_2' ? 'Y' : '',
          PaytypeExtendedWarranty: this.Paytype[3] == 'Extended_3' ? 'Y' : '',

          AgeFrom: this.AgeFrom,
          AgeTo: this.AgeTo,
          dataLength: this.ServiceData.length,
          inventory: this.inventory == 'All' ? '' : this.inventory,
          ROSTATUS: this.rostatus == 'All' ? '' : this.rostatus.toString(),
          topfive: 'Y',
          data: Item,
          Total: Item.RoInfo ? Item.RoInfo.length : 0
        },
      ];
    } else {
      this.RODetailsObjectMain = [
        {
          StartDate: this.FromDate,
          EndDate: this.ToDate,
          var1: this.selectedDataGrouping[0]?.columnname,
          var2: this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '',
          // var3: this.path3,
          var1Value: ParentItem.data1,
          var2Value: Item.data2,
          GrossTypeLabor: this.Department.indexOf('Service') >= 0 ? 'Y' : '',
          GrossTypeParts: this.Department.indexOf('Parts') >= 0 ? 'Y' : '',
          GrossTypeMisc: this.Department.indexOf('Misc') >= 0 ? 'Y' : '',
          GrossTypeSublet: this.Department.indexOf('Sublet') >= 0 ? 'Y' : '',
          PaytypeCP: this.Paytype[0] == 'Customerpay_0' ? 'Y' : '',
          PaytypeWarranty: this.Paytype[1] == 'Warranty_1' ? 'Y' : '',
          PaytypeInternal: this.Paytype[2] == 'Internal_2' ? 'Y' : '',
          PaytypeExtendedWarranty: this.Paytype[3] == 'Extended_3' ? 'Y' : '',

          AgeFrom: this.AgeFrom,
          AgeTo: this.AgeTo,
          dataLength: this.ServiceData.length,
          inventory: this.inventory == 'All' ? '' : this.inventory,
          ROSTATUS: this.rostatus == 'All' ? '' : this.rostatus.toString(),
          topfive: '',
          Total: Item.Repair_Orders
          // var3Value: Item.data3,
          // userName: Item.data3,
        },
      ];
    }
  }
  isDesc: boolean = false;
  column: string = 'CategoryName';
  sort(property: any, data: any) {
    this.isDesc = !this.isDesc; //change the direction
    this.column = property;
    let direction = this.isDesc ? 1 : -1;
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
  RODetailsObjectMain: any = []
  scrollpositionstoring: any;
  ngAfterViewInit() {
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.comm.pageName == 'Service Open RO') {
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
        if (res.obj.title == 'Service Open RO') {
          if (res.obj.state == true) {
            this.exportToExcel();
          }
        }
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res) => {
      if (this.print != undefined) {
        if (res.obj.title == 'Service Open RO') {
          if (res.obj.statePrint == true) {
            //   this.GetPrintData();
          }
        }
      }
    });
    this.Pdf = this.shared.api.getExportToPDFAllReports().subscribe((res) => {
      if (this.Pdf != undefined) {
        if (res.obj.title == 'Service Open RO') {
          if (res.obj.statePDF == true) {
            //     this.generatePDF();
          }
        }
      }
    });
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res) => {
      if (this.email != undefined) {
        if (res.obj.title == 'Service Open RO') {
          if (res.obj.stateEmailPdf == true) {
            //    this.sendEmailData(res.obj.Email, res.obj.notes, res.obj.from);
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
  paytypeDisplay: any = []
  multipleorsingle(block: any, e: any) {

    if (block == 'Dept') {
      const index = this.Department.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.Department.splice(index, 1);
      } else {
        this.Department.push(e);
      }

    }

    if (block == 'TB') {
      this.TotalReport = e;
    }
    if (block == 'I') {
      this.inventory = [];
      this.inventory.push(e);

    }
    if (block == 'RO') {
      if (e == 'All') {
        const index = this.rostatus.findIndex((i: any) => i == e);
        if (index >= 0) {
          this.rostatus.splice(index, 1);
        } else {
          this.rostatus = []
          this.rostatus.push(e);
        }
      }
      else {
        const Allindex = this.rostatus.findIndex((i: any) => i == 'All');
        if (Allindex >= 0) {
          this.rostatus = []
        }
        const index = this.rostatus.findIndex((i: any) => i == e);
        if (index >= 0) {
          this.rostatus.splice(index, 1);
        } else {
          this.rostatus.push(e);
        }
      }

    }


    if (block == 'PT') {
      let spliting = e.split('_');
      const index = this.Paytype.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.Paytype[spliting[1]] = '';
        let arr = this.Paytype.filter((e: any) => e != '');
        if (arr.length == 0) {
          this.toast.show('Please select atleast one Pay Type', 'warning', 'Warning');
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
          this.toast.show('Please select atleast one Gross Type', 'warning', 'Warning');
        }
      } else {
        this.Grosstype[spliting[1]] = e;
      }

    }
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
  activePopover: number = -1;

  togglePopover(popoverIndex: number) {
    this.activePopover = this.activePopover === popoverIndex ? -1 : popoverIndex;
  }

  viewreport() {
    this.activePopover = -1

    if (this.selectedDataGrouping.length == 0) {
      this.toast.show('Please select atleast one Value from Grouping', 'warning', 'Warning');
    } else {
      if (this.storeIds.length == 0 && this.selectedotherstoreids.length == 0) {
        this.toast.show('Please select atleast one Store', 'warning', 'Warning');
      } else {
        let gt = this.Grosstype.filter((e: any) => e != '');
        let pt = this.Paytype.filter((e: any) => e != '');
        if (pt.length == 0 || gt.length == 0) {
          if (pt.length == 0) {
            this.toast.show('Please select atleast one Pay Type', 'warning', 'Warning');
          }
          if (gt.length == 0) {
            this.toast.show('Please select atleast one Gross Type', 'warning', 'Warning');
          }
        } else {
          this.setHeaderData()

          this.getServiceData()
        }
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
    const ServiceData = this.ServiceData.map((_arrayElement: any) =>
      Object.assign({}, _arrayElement)
    );
    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('Service Open RO');
    worksheet.views = [
      {
        showGridLines: false,
      },
    ];
    worksheet.addRow('');
    const titleRow = worksheet.addRow(['Service Open RO']);
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
    groupings.value = this.selectedDataGrouping[0]?.ARG_LABEL + ', ' + this.selectedDataGrouping[1]?.ARG_LABEL
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
    Groups.value = 'Group :';
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
    stores1.value = this.ExcelStoreNames == 0
      ? 'All Stores'
      : this.ExcelStoreNames == null
        ? '-'
        : this.ExcelStoreNames.toString().replaceAll(',', ', ');
    stores1.font = { name: 'Arial', family: 4, size: 9 };
    stores1.alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    const Filters = worksheet.addRow(['Filters :']);
    Filters.getCell(1).font = { name: 'Arial', family: 4, size: 9, bold: true };
    // const NewUsed = worksheet.getCell('B12');
    // NewUsed.value = 'RO Type :';
    // const newused = worksheet.getCell('D12');
    // newused.value = '-';
    // newused.font = { name: 'Arial', family: 4, size: 9 };
    // const DealType = worksheet.getCell('B13');
    // DealType.value = 'Department :';
    // const dealtype = worksheet.getCell('D13');
    // dealtype.value = '-';
    // dealtype.font = { name: 'Arial', family: 4, size: 9 };
    // const Source = worksheet.getCell('B15');
    // Source.value = 'Select Target :';
    // const source = worksheet.getCell('D15');
    // source.value = '-';
    // source.font = { name: 'Arial', family: 4, size: 9 };
    // const Chargebacks = worksheet.getCell('B16');
    // Chargebacks.value = 'Gross Type :';
    // const chargebacks = worksheet.getCell('D16');
    // chargebacks.value =
    //   this.Grosstype[0] == 'Labour_0' ? 'Labour' : '' +
    //     ',' +
    //     this.Grosstype[1] == 'Parts_1' ? 'Parts' : '' +
    //       ',' +
    //       this.Grosstype[2] == 'Misc_2' ? 'Misc' : '' +
    //         ',' +
    //         this.Grosstype[3] == 'Sublet_3' ? 'Sublet' : '';
    // chargebacks.font = { name: 'Arial', family: 4, size: 9 };
    // const PackHoldback = worksheet.getCell('B17');
    // PackHoldback.value = 'Source :';
    // const packholdback = worksheet.getCell('D17');
    // packholdback.value = '-';
    // packholdback.font = { name: 'Arial', family: 4, size: 9 };
    // const ROType = worksheet.getCell('B11');
    // ROType.value = 'RO Type :';
    // const rotype = worksheet.getCell('D11');
    // rotype.value = this.ROType == ''
    // ? '-'
    // : this.ROType == null
    // ? '-'
    // : this.ROType.toString().replaceAll(',', ', ');
    // rotype.font = { name: 'Arial', family: 4, size: 9 };
    const Department = worksheet.getCell('B13');
    Department.value = 'Department :';
    const department = worksheet.getCell('D13');
    department.value = this.Department == ''
      ? '-'
      : this.Department == null
        ? '-'
        : this.Department.toString().replaceAll(',', ', ');;
    department.font = { name: 'Arial', family: 4, size: 9 };
    const PayType = worksheet.getCell('B14');
    PayType.value = 'Pay Type :';
    const paytype = worksheet.getCell('D14');
    let paytypeformat = this.Paytype.toString().replace(/[0-9_]/g, '')
    paytype.value = paytypeformat.replaceAll(',', ',  ')
    paytypeformat == ''
      ? '-'
      : paytypeformat == null
        ? '-'
        : paytypeformat.replaceAll(',', ',  ');
    paytype.font = { name: 'Arial', family: 4, size: 9 };
    const SelectTarget = worksheet.getCell('B14');
    SelectTarget.value = "Inventory RO's :";
    const selecttarget = worksheet.getCell('D14');
    selecttarget.value = this.inventory == 'All'
      ? 'All'
      : 'Inventory'
    selecttarget.font = { name: 'Arial', family: 4, size: 9 };
    const DealStatus = worksheet.getCell('B15');
    DealStatus.value = 'Pay Type :';
    const dealstatus = worksheet.getCell('D15');
    let deal = this.Paytype.toString().replace(/[0-9_]/g, '')
    dealstatus.value = deal
    // (this.Paytype[0] == 'Customerpay_0' ? 'Customer Pay, ' : '' )+  (this.Paytype[1] == 'Warranty_1' ? 'Warranty, ' : '') + (this.Paytype[2] == 'Internal_2' ? 'Internal' : '');
    dealstatus.font = { name: 'Arial', family: 4, size: 9 };
    // const GrossType = worksheet.getCell('B15');
    // GrossType.value = 'Gross Type :';
    // const grosstype = worksheet.getCell('D15');
    // grosstype.value =
    // this.Grosstype == ''
    // ? '-'
    // : this.Grosstype == null
    // ? '-'
    // : this.Grosstype.toString().replaceAll(',', ', ');
    // grosstype.font = { name: 'Arial', family: 4, size: 9 };
    // const Source = worksheet.getCell('B16');
    // Source.value = 'Source :';
    // const source = worksheet.getCell('D16');
    // source.value =this.Transactorgl == 'T' ? 'Transaction' : (this.Transactorgl == 'G' ? 'GL' : '');
    // source.font = { name: 'Arial', family: 4, size: 9 };
    worksheet.addRow('');
    let dateYear = worksheet.getCell('A16');
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
    worksheet.mergeCells('B16', 'D16');
    let units = worksheet.getCell('B16');
    units.value = 'Open RO';
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
    units.border = { right: { style: 'thin' } };
    worksheet.mergeCells('E16', 'H16');
    let frontgross = worksheet.getCell('H16');
    frontgross.value = 'Labor';
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
    frontgross.border = { right: { style: 'thin' } };
    worksheet.mergeCells('I16', 'L16');
    let backgross = worksheet.getCell('L16');
    backgross.value = 'Parts';
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
    worksheet.mergeCells('M16', 'P16');
    let totalgross = worksheet.getCell('P16');
    totalgross.value = '';
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
    //console.log(FromDate, ToDate, '........From Date and To Date');
    let Headings = [
      ((FromDate == '' || FromDate == null) && (ToDate == '' || ToDate == null)) ? "All Open RO's" : FromDate + ' - ' + ToDate + ', ' + PresentYear,
      'Qty',
      'Guide',
      'Hours',
      'Labor Sales',
      'Labor Cost',
      'Discounts',
      'Labor Gross',
      'Parts Sales',
      'Parts Cost',
      'Discounts',
      'Parts Gross',
      'Total Sales',
      'Total Gross',
      'Total Gross PR',
      'PR GP',
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
    for (const d of ServiceData) {
      var obj = [
        d.data1 == '' ? '-' : d.data1 == null ? '-' : d.data1,
        d.Repair_Orders == '' ? '-' : d.Repair_Orders == null ? '-' : d.Repair_Orders,
        d.Guide == '' ? '-' : d.Guide == null ? '-' : d.Guide,
        d.Total_Hours == '' ? '-' : d.Total_Hours == null ? '-' : d.Total_Hours,
        d.laborsale == '' ? '-' : d.laborsale == null ? '-' : d.laborsale,
        d.laborcost == '' ? '-' : d.laborcost == null ? '-' : d.laborcost,
        d.ServiceDiscount == '' ? '-' : d.ServiceDiscount == null ? '-' : d.ServiceDiscount,
        d.laborgross == '' ? '-' : d.laborgross == null ? '-' : d.laborgross,
        d.PartsSale == '' ? '-' : d.PartsSale == null ? '-' : d.PartsSale,
        d.PartsCost == '' ? '-' : d.PartsCost == null ? '-' : d.PartsCost,
        d.partsDiscount == '' ? '-' : d.partsDiscount == null ? '-' : d.partsDiscount,
        d.partsgross == '' ? '-' : d.partsgross == null ? '-' : d.partsgross,
        d.TotalSale == '' ? '-' : d.TotalSale == null ? '-' : d.TotalSale,
        d.Total_Gross == '' ? '-' : d.Total_Gross == null ? '-' : d.Total_Gross,
        d.Total_Gross_PR == '' ? '-' : d.Total_Gross_PR == null ? '-' : d.Total_Gross_PR,
        d.PR_GP == '' ? '-' : d.PR_GP == null ? '-' : d.PR_GP + '%',



      ]
      const Data1 = worksheet.addRow(obj);
      // Data1.outlineLevel = 1; // Grouping level 1
      Data1.font = { name: 'Arial', family: 4, size: 9 };
      Data1.getCell(1).alignment = {
        indent: 1,
        vertical: 'middle',
        horizontal: 'left',
      };
      Data1.eachCell((cell: any, number: any) => {
        cell.border = { right: { style: 'thin' } };
        if (
          // (number > 1 && number < 7) ||
          number == 2
        ) {
          cell.numFmt = '#,##0';
        } if (number == 4) {
          cell.numFmt = '#,##0.00';
        } if (number > 4) {
          cell.numFmt = '$#,##0';
        }
        if (number > 1 && obj[number] != undefined) {
          // cell.alignment = { horizontal: 'right', vertical: 'middle', indent: 1 };
          if (obj[number] < 0) {
            Data1.getCell(number + 1).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } };
          }
        }
        if (number == 1) {
          if (obj[number] < 0) {
            Data1.getCell(number + 1).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' }, bold: true };
          }
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
      // if (d.data1 === 'REPORTS TOTAL') {
      //   Data1.eachCell((cell, number: any) => {
      //     cell.font = { name: 'Arial', family: 4, size: 9, bold: true };
      //     if (number > 1 && obj[number] != undefined) {
      //       if (obj[number] < 0) {
      //         Data1.getCell(number + 1).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } };
      //       }
      //     }
      //   });
      // }
      if (d.Data2 != undefined) {
        //console.log('Hellooo..................');
        for (const d1 of d.Data2) {
          var obj = [
            d1.data2 == '' ? '-' : d1.data2 == null ? '-' : d1.data2,
            d1.Repair_Orders == '' ? '-' : d1.Repair_Orders == null ? '-' : d1.Repair_Orders,
            d.Guide == '' ? '-' : d.Guide == null ? '-' : d.Guide,

            d1.Total_Hours == '' ? '-' : d1.Total_Hours == null ? '-' : d1.Total_Hours,
            d1.laborsale == '' ? '-' : d1.laborsale == null ? '-' : d1.laborsale,
            d1.laborcost == '' ? '-' : d1.laborcost == null ? '-' : d1.laborcost,
            d1.ServiceDiscount == '' ? '-' : d1.ServiceDiscount == null ? '-' : d1.ServiceDiscount,
            d1.laborgross == '' ? '-' : d1.laborgross == null ? '-' : d1.laborgross,
            d1.PartsSale == '' ? '-' : d1.PartsSale == null ? '-' : d1.PartsSale,
            d1.PartsCost == '' ? '-' : d1.PartsCost == null ? '-' : d1.PartsCost,
            d1.partsDiscount == '' ? '-' : d1.partsDiscount == null ? '-' : d1.partsDiscount,
            d1.partsgross == '' ? '-' : d1.partsgross == null ? '-' : d1.partsgross,
            d1.TotalSale == '' ? '-' : d1.TotalSale == null ? '-' : d1.TotalSale,
            d1.Total_Gross == '' ? '-' : d1.Total_Gross == null ? '-' : d1.Total_Gross,
            d1.Total_Gross_PR == '' ? '-' : d1.Total_Gross_PR == null ? '-' : d1.Total_Gross_PR,
            d1.PR_GP == '' ? '-' : d1.PR_GP == null ? '-' : d1.PR_GP + '%',
          ]
          const Data2 = worksheet.addRow(obj);
          Data2.outlineLevel = 1; // Grouping level 2
          Data2.font = { name: 'Arial', family: 4, size: 8 };
          Data2.getCell(1).alignment = {
            indent: 2,
            vertical: 'middle',
            horizontal: 'left',
          };
          Data2.eachCell((cell: any, number: any) => {
            cell.border = { right: { style: 'thin' } };
            if (
              // (number > 1 && number < 7) ||
              number == 2
            ) {
              cell.numFmt = '#,##0';
            } if (number == 4) {
              cell.numFmt = '#,##0.00';
            } if (number > 4) {
              cell.numFmt = '$#,##0';
            }
            if (number > 1 && obj[number] != undefined) {
              // cell.alignment = { horizontal: 'right', vertical: 'middle', indent: 2 };
              if (obj[number] < 0) {
                Data2.getCell(number + 1).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } };
              }
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
        }
      }
    }
    worksheet.eachRow((row, rowIndex) => {
      row.eachCell((cell, colIndex) => {
        if (rowIndex > 1 && rowIndex < 15) { // Skip the header row
          // Apply conditional alignment based on your conditions
          if (colIndex === 1) {
            // Apply right alignment to the second column
            cell.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
          }
        }
      });
    });
    // worksheet.getColumn(1).alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
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

    this.shared.exportToExcel(workbook, 'Service Open RO');

  }

}