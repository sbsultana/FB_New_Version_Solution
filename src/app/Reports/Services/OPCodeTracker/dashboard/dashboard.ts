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
import { OPCodeTrackerDetails } from '../opcode-tracker-details/opcode-tracker-details';
import { NgbModule } from '@ng-bootstrap/ng-bootstrap';


@Component({
  selector: 'app-dashboard',
  imports: [SharedModule, BsDatepickerModule, DateRangePicker, Stores, NgbModule],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
})
export class Dashboard {
  ServiceData: any = [];
  IndividualServiceGross: any = [];
  TotalServiceGross: any = [];


  TotalReport: string = 'T';
  NoData: boolean = false;
  ROType: any = 'Closed';
  Department: any = ['Service'];
  Paytype: any = ['Customerpay_0', 'Warranty_1', 'Internal_2', 'Extended_3'];
  Grosstype: any = ['Labour_0', '', 'Misc_2', 'Sublet_3'];
  zeroHours: any = ['Y']
  labourType: any = 'N';
  Target: any = 'F';

  responcestatus: any = '';
  CurrentDate = new Date();

  topfive: boolean = false;

  actionType: any = 'N'
  reportOpenSub!: Subscription;
  reportGetting!: Subscription;
  Pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;
  SearchText: any = '';

  stores: any = []
  groupsArray: any = [];
  storename: any = ''
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  groupId: any = 0;
  storeIds: any = 0;
  DefaultLoad: any = 'E'

  storesFilterData: any = {
    'groupsArray': this.groupsArray, 'groupId': this.groupId, 'storesArray': this.stores, 'storeids': '1', 'type': 'M', 'others': 'N',
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
  dataGrouping = [
    { id: 40, columnname: 'Store_Name', Active: 'Y', ARG_LABEL: 'Store' },
    { id: 26, columnname: 'ServiceAdvisor_Name', Active: 'Y', ARG_LABEL: 'Advisor Name' },
    { id: 27, columnname: 'ServiceAdvisor', Active: 'Y', ARG_LABEL: 'Advisor Number' },
    { id: 28, columnname: 'techname', Active: 'Y', ARG_LABEL: 'Tech Name' },
    { id: 29, columnname: 'techno', Active: 'Y', ARG_LABEL: 'Tech Number' },
    { id: 69, columnname: 'opcode', Active: 'Y', ARG_LABEL: 'Op Code' },
    { id: 70, columnname: 'ronumber', Active: 'Y', ARG_LABEL: 'RO Number' },
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
    this.selectedDataGrouping.push(this.dataGrouping[5])

    this.shared.setTitle(this.comm.titleName + '-OP Code Tracker');
    this.initializeDates('MTD')
    if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
      if (localStorage.getItem('flag') == 'V') {
        this.Department = ['Service'];
      } else {
        this.Department = ['Service', 'Parts', 'Body']
      }
      this.groupId = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Preferences
      this.storeIds = ''
    }
    if (this.shared.common.groupsandstores.length > 0) {
      this.groupsArray = this.shared.common.groupsandstores.filter((val: any) => val.sg_id != this.shared.common.reconID);
      this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
      this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_Name : this.groupName = ''
      this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
      this.getStoresandGroupsValues()
    }
    this.setHeaderData()

  }

  setHeaderData() {
    const data = {
      title: 'OP Code Tracker',
      stores: this.storeIds,
      ROType: this.ROType,
      Department: this.Department,
      Paytype: this.Paytype,
      Target: this.Target,
      ToporBottom: this.TotalReport,
      fromdate: this.FromDate,
      todate: this.ToDate,
      grosstype: this.Grosstype,
      groups: this.groupId,
      labourtype: this.labourType,
      zeroHours: this.zeroHours,
      searchtext: this.SearchText,
      dataGroupings: this.selectedDataGrouping,

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
  getServiceData() {
    if (this.storeIds != '') {
      this.responcestatus = '';
      this.shared.spinner.show();
      this.GetData();
    } else {
      this.NoData = true
    }

  }

  filterGrid(e: any) {
    console.log(e, '.............');

    // this.SearchText = e.target.value
    if (this.CompleteServiceData != undefined && this.CompleteServiceData.length > 0) {
      if (this.SearchText.trim() !== '') {
        if (this.selectedDataGrouping[0]?.columnname == 'opcode') {
          this.ServiceData = this.CompleteServiceData.filter((item: any) =>
            (item.OpcodeDesc && item.OpcodeDesc.toLowerCase().includes(this.SearchText.toLowerCase())) ||
            (item.data1 && item.data1.toLowerCase().includes(this.SearchText.toLowerCase()))
          );
          if (this.ServiceData.length == 0) {
            this.NoData = true
          } else {
            this.NoData = false
          }
        }
        if (this.selectedDataGrouping[1]?.columnname == 'opcode') {
          console.log(this.SearchText);
          this.ServiceData = this.CompleteServiceData
          let data = this.CompleteServiceData.some((x: any) => {
            console.log(this.SearchText);
            if (x.Data2 != undefined) {
              x.Data2 = x.Duplicate.filter((item: any) =>
                (item.OpcodeDesc && item.OpcodeDesc.toLowerCase().includes(this.SearchText.toLowerCase())) ||
                (item.data2 && item.data2.toLowerCase().includes(this.SearchText.toLowerCase()))
              );
            }
          });
        }
        if (this.selectedDataGrouping[2]?.columnname == 'opcode') {
          console.log(this.SearchText);
          this.ServiceData = this.CompleteServiceData
          let data = this.CompleteServiceData.some((x: any) => {
            console.log(this.SearchText);
            if (x.Data2 != undefined) {
              x.Data2.forEach((val: any) => {
                if (val.DuplicateSubData != undefined) {
                  val.SubData = val.DuplicateSubData.filter((item: any) =>
                    (item.OpcodeDesc && item.OpcodeDesc.toLowerCase().includes(this.SearchText.toLowerCase())) ||
                    (item.data3 && item.data3.toLowerCase().includes(this.SearchText.toLowerCase()))
                  );
                }
              })
            }
          });
        }
      } else {
        this.NoData = false
        this.ServiceData = [...this.CompleteServiceData]
      }
    }


    //   this.TitleData = this.ServiceData.filter((item: any) =>
    //     (item.data1 && item.data1.toLowerCase().includes(this.SearchText.toLowerCase())));
    // } else {
    //   this.TitleData = [...this.TitleCompleteData];
    // }
    // this.TitleData.length > 0 ? this.NoData = false : this.NoData = true;
    // this.callLoadingState == 'ANS' ? this.sort(this.column, this.TitleData, this.callLoadingState) : ''
    // let position = this.scrollpositionstoring + 10
    // setTimeout(() => {
    //   this.scrollcent.nativeElement.scrollTop = position
    // }, 500);
  }
  key: any = 'Count'
  order: any = 'asc'
  sort(key: any, order: any) {
    if (this.key == key) {
      if (order == 'asc') {
        this.order = 'desc'
      } else {
        this.order = 'asc'
      }
    } else {
      this.order = 'asc'
    }
    this.key = key
    this.getServiceData()
  }
  CompleteServiceData: any = []
  GetData() {
    console.log(this.selectedDataGrouping, 'Selected Data Grouping');

    this.IndividualServiceGross = [];
    const obj = {
      startdate: this.FromDate,
      enddate: this.ToDate,
      StoreID: this.storeIds == 0 ? '' : this.storeIds.toString(),
      var1: this.selectedDataGrouping.length >= 1 ? this.selectedDataGrouping[0]?.columnname : '',
      var2: this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '',
      var3: this.selectedDataGrouping.length == 3 ? this.selectedDataGrouping[2]?.columnname : '',
      excludeZeroHours: this.zeroHours.toString() == 'N' ? '' : this.zeroHours.toString(),
      type: 'D',
      topOpcodes: this.topfive == true ? 'Y' : '',
      sortExp: this.key + ' ' + this.order,
      searchtext: this.SearchText,
      paytype: this.Paytype.map((item: any) => item.charAt(0).toUpperCase()).filter((item: any) => item !== '').toString()
    };
    const curl = environment.apiUrl + 'fredbeans/GetServiceOPCodeTrakerPaytype';
    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetServiceOPCodeTrakerPaytype', obj).subscribe(
      (res) => {
        const currentTitle = document.title;
        this.shared.api.logSaving(curl, {}, '', res.message, currentTitle);
        if (res.status == 200) {
          if (res.response != undefined) {
            if (res.response.length > 0) {
              this.IndividualServiceGross = res.response;

              this.responcestatus = this.responcestatus + 'I';
              this.GetTotalData();
              this.NoData = false;
              let path2 = this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '';
              let topfive = this.topfive;
              this.IndividualServiceGross.some(function (x: any) {
                if (x.Data2 != undefined) {
                  x.Data2 = JSON.parse(x.Data2);
                  x.Data2 = x.Data2.map((v: any) => ({
                    ...v,
                    SubData: [],
                    DuplicateSubData: [],
                    data2sign: '+',
                  }));
                  x.Duplicate = x.Data2
                }
                if (x.Data3 != undefined) {
                  x.Data3 = JSON.parse(x.Data3);
                  x.Data2.forEach((val: any) => {
                    x.Data3.forEach((ele: any) => {
                      if (val.data2 == ele.data2) {
                        val.SubData.push(ele);
                        val.DuplicateSubData.push(ele);
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
      StoreID: this.storeIds == 0 ? '' : this.storeIds.toString(),

      var1: this.selectedDataGrouping.length >= 1 ? this.selectedDataGrouping[0]?.columnname : '',
      var2: this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '',
      var3: this.selectedDataGrouping.length == 3 ? this.selectedDataGrouping[2]?.columnname : '',
      excludeZeroHours: this.zeroHours.toString() == 'N' ? '' : this.zeroHours.toString(),
      type: 'T',
      topOpcodes: this.topfive == true ? 'Y' : '',
      sortExp: this.key + ' ' + this.order,
      searchtext: this.SearchText,
      paytype: this.Paytype.map((item: any) => item.charAt(0).toUpperCase()).filter((item: any) => item !== '').toString()

    };
    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetServiceOPCodeTrakerPaytype', obj).subscribe(
      (totalres) => {
        if (totalres.status == 200) {
          if (totalres.response != undefined) {
            if (totalres.response.length > 0) {
              //   this.TotalServiceGross = [
              //     {
              //         "OpcodeDesc": "",
              //         "Count": 12011,
              //         "Mileage": 0,
              //         "Hours": 6284.45,
              //         "LaborSale": 1025200.68,
              //         "LaborGross": 806239.91,
              //         "LaborGP": 78.64,
              //         "ELR": 163.13,
              //         "avg_hrs": 0.523224,
              //         "avg_LaborSale": 85.355147,
              //         "avg_LaborGP": 67.125127,
              //         "PartsSale": 852439.14,
              //         "PartsGross": 358476.73,
              //         "PartsGP": 42.05,
              //         "avg_PartsSale": 70.971537,
              //         "avg_PartsGP": 29.845702
              //     }
              // ]
              // this.TotalServiceGross = this.TotalServiceGross.map((v: any) => ({
              //   ...v,
              //   data1: 'REPORTS TOTAL',
              //   Dealer: '+',
              //   Data2: [],
              //   Duplicate:[],
              // }));
              this.TotalServiceGross = totalres.response.map((v: any) => ({
                ...v,
                data1: 'REPORTS TOTAL',
                Dealer: '+',
                Data2: [],
                Duplicate: [],
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
        this.CompleteServiceData = this.ServiceData;
        // if (this.topfive == true) {
        //   setTimeout(() => {
        //     this.ServiceData.forEach((val: any, i: any) => {
        //       if (i < this.ServiceData.length - 1) {
        //         setTimeout(() => {
        //           val.Dealer = '-'
        //           if (this.path2 == '') {
        //             this.openRODetails(val);
        //           } else {
        //             val.Data2.forEach((sub: any, j: any) => {
        //               sub.data2sign = '-'
        //               this.openRODetails(sub, val);
        //             })
        //           }
        //         }, 100);
        //       }
        //     })
        //   }, 300);
        // }
      } else {
        this.IndividualServiceGross.unshift(this.TotalServiceGross[0]);
        this.ServiceData = this.IndividualServiceGross;
        this.CompleteServiceData = this.ServiceData

      }
      this.shared.spinner.hide();
    } else if (this.responcestatus == 'T') {
      this.ServiceData = this.TotalServiceGross;
      this.CompleteServiceData = this.ServiceData

    } else if (this.responcestatus == 'I') {
      this.ServiceData = this.IndividualServiceGross;
      this.CompleteServiceData = this.ServiceData

    } else {
      this.NoData = true;
    }
    console.log(this.ServiceData);

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
    if (id == 'D_' + ind) {
      if ((this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '') == '') {
        // console.log(this.path2);
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
      // // //console.log(this.path3)
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
  // isDesc: boolean = false;
  // column: string = 'CategoryName';

  // sort(property: any, data: any) {
  //   this.isDesc = !this.isDesc; //change the direction
  //   this.column = property;
  //   let direction = this.isDesc ? 1 : -1;
  //   // console.log(property)
  //   data.sort(function (a: any, b: any) {
  //     if (a[property] < b[property]) {
  //       return -1 * direction;
  //     } else if (a[property] > b[property]) {
  //       return 1 * direction;
  //     } else {
  //       return 0;
  //     }
  //   });
  // }
  openDetails(Item: any, ParentItem: any, ref: any) {
    // console.log(Item, ParentItem);
    if (ref == '1') {
      if (Item.data1 != undefined && Item.data1 != 'Reports Total') {
        const DetailsServicePerson = this.shared.ngbmodal.open(
          OPCodeTrackerDetails,
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
            DepartmentS: this.Department.indexOf('Service') >= 0 ? 'S' : '',
            DepartmentP: this.Department.indexOf('Parts') >= 0 ? 'P' : '',            // DepartmentP: '',
            DepartmentQ: this.Department.indexOf('Quicklube') >= 0 ? 'Q' : '',
            DepartmentB: this.Department.indexOf('Body') >= 0 ? 'B' : '',
            PolicyAccount: this.labourType,
            userName: Item.data1,
            Grosstype: this.Grosstype,
            layer: 1,
            zeroHours: this.zeroHours.toString() == 'N' ? '' : this.zeroHours.toString(),
            topFive: this.topfive == true ? 'Y' : '',
            paytype: this.Paytype.map((item: any) => item.charAt(0).toUpperCase()).filter((item: any) => item !== '').toString()

          },
        ];
      }
    }
    // console.log(Item)
    if (ref == '2') {
      if (Item.data2 != undefined) {
        const DetailsServicePerson = this.shared.ngbmodal.open(
          OPCodeTrackerDetails,
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
            DepartmentS: this.Department.indexOf('Service') >= 0 ? 'S' : '',
            DepartmentP: this.Department.indexOf('Parts') >= 0 ? 'P' : '',            // DepartmentP: '',
            DepartmentQ: this.Department.indexOf('Quicklube') >= 0 ? 'Q' : '',
            DepartmentB: this.Department.indexOf('Body') >= 0 ? 'B' : '',
            PolicyAccount: this.labourType,
            userName: Item.data2,
            Grosstype: this.Grosstype,
            layer: 2,
            zeroHours: this.zeroHours.toString() == 'N' ? '' : this.zeroHours.toString(),
            topFive: this.topfive == true ? 'Y' : '',
            paytype: this.Paytype.map((item: any) => item.charAt(0).toUpperCase()).filter((item: any) => item !== '').toString()


          },
        ];
      }
    }

    if (ref == '3') {
      if (Item.data3 != undefined) {
        const DetailsServicePerson = this.shared.ngbmodal.open(
          OPCodeTrackerDetails,
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
            DepartmentS: this.Department.indexOf('Service') >= 0 ? 'S' : '',
            DepartmentP: this.Department.indexOf('Parts') >= 0 ? 'P' : '',            // DepartmentP: '',
            DepartmentQ: this.Department.indexOf('Quicklube') >= 0 ? 'Q' : '',
            DepartmentB: this.Department.indexOf('Body') >= 0 ? 'B' : '',
            PolicyAccount: this.labourType,
            userName: Item.data3,
            Grosstype: this.Grosstype,
            layer: 3,
            zeroHours: this.zeroHours.toString() == 'N' ? '' : this.zeroHours.toString(),
            topFive: this.topfive == true ? 'Y' : '',
            paytype: this.Paytype.map((item: any) => item.charAt(0).toUpperCase()).filter((item: any) => item !== '').toString()


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

  }


  ngAfterViewInit() {

    this.shared.api.getStores().subscribe((res: any) => {
      if (this.comm.pageName == 'OP Code Tracker') {
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
        if (res.obj.title == 'OP Code Tracker') {
          if (res.obj.state == true) {
            this.exportToExcel();
          }
        }
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res) => {
      if (this.print != undefined) {
        if (res.obj.title == 'OP Code Tracker') {
          if (res.obj.statePrint == true) {
            //this.GetPrintData();
          }
        }
      }
    });


    this.Pdf = this.shared.api.getExportToPDFAllReports().subscribe((res) => {
      if (this.Pdf != undefined) {
        if (res.obj.title == 'OP Code Tracker') {
          if (res.obj.statePDF == true) {
            //    this.generatePDF();
          }
        }
      }
    });
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res) => {
      if (this.email != undefined) {
        if (res.obj.title == 'OP Code Tracker') {
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
      'type': 'M', 'others': 'N', 'DefaultLoad': this.DefaultLoad
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

  multipleorsingle(block: any, e: any) {

    if (block == 'ZH') {
      this.zeroHours = [];
      this.zeroHours.push(e)
    }

    if (block == 'TB') {
      this.TotalReport = e;

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
      if (this.storeIds.length == 0) {
        this.toast.show('Please select atleast one Store', 'warning', 'Warning');
      } else {
        let pt = this.Paytype.filter((e: any) => e != 'warning', 'Warning');
        if (pt.length == 0) {
          this.toast.show('Please select atleast one Pay Type', 'warning', 'Warning');
        }
        else {
          this.actionType = 'Y';
          this.DefaultLoad = ''
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
    // const obj = {
    //   id: this.groups,
    //   userid: JSON.parse(localStorage.getItem('UserDetails')!).userid,
    // };
    // this.shared.api
    //   .postmethodOne('fredbeans/GetStoresbyGroupuserid', obj)
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
    const worksheet = workbook.addWorksheet('OP Code Tracker');
    worksheet.views = [
      {
        showGridLines: false,
      },
    ];
    worksheet.addRow('');
    const titleRow = worksheet.addRow(['OP Code Tracker']);
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
    // const ROType = worksheet.getCell('B11');
    // ROType.value = 'RO Type :';
    // const rotype = worksheet.getCell('D11');
    // rotype.value = '';
    // rotype.font = { name: 'Arial', family: 4, size: 9 };
    // const Department = worksheet.getCell('B12');
    // Department.value = 'Department :';
    // const department = worksheet.getCell('D12');
    // department.value = '';
    // department.font = { name: 'Arial', family: 4, size: 9 };
    // const PayType = worksheet.getCell('B13');
    // PayType.value = 'Pay Type :';
    // const paytype = worksheet.getCell('D13');
    // paytype.value =
    //   this.Paytype[0] + ',' + this.Paytype[1] + ',' + this.Paytype[2];
    // paytype.font = { name: 'Arial', family: 4, size: 9 };
    // const SelectTarget = worksheet.getCell('B14');
    // SelectTarget.value = 'Select Target :';
    // const selecttarget = worksheet.getCell('D14');
    // selecttarget.value = '';
    // selecttarget.font = { name: 'Arial', family: 4, size: 9 };
    // const GrossType = worksheet.getCell('B15');
    // GrossType.value = 'Gross Type :';
    // const grosstype = worksheet.getCell('D15');
    // grosstype.value =
    //   this.Grosstype[0] +
    //   ',' +
    //   this.Grosstype[1] +
    //   ',' +
    //   this.Grosstype[2] +
    //   ',' +
    //   this.Grosstype[3];
    // grosstype.font = { name: 'Arial', family: 4, size: 9 };
    // const Source = worksheet.getCell('B16');
    // Source.value = 'Source :';
    // const source = worksheet.getCell('D16');
    // source.value = '';
    // source.font = { name: 'Arial', family: 4, size: 9 };
    // const ReportTotals = worksheet.getCell('B17');
    // ReportTotals.value = 'Report Totals :';
    // const reporttotals = worksheet.getCell('D17');
    // reporttotals.value = this.TotalReport == 'T' ? 'Top' : 'Bottom';
    // reporttotals.font = { name: 'Arial', family: 4, size: 9 };
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


    let Desc = worksheet.getCell('B19');
    Desc.value = '';
    Desc.alignment = { vertical: 'middle', horizontal: 'center' };
    Desc.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
      color: { argb: 'FFFFFF' },
    };
    Desc.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '2a91f0' },
      bgColor: { argb: 'FF0000FF' },
    };
    Desc.border = { right: { style: 'dotted' } };


    worksheet.mergeCells('C19', 'D19');
    let frontgross = worksheet.getCell('C19');
    frontgross.value = 'Total';
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

    worksheet.mergeCells('E19', 'K19');
    let backgross = worksheet.getCell('E19');
    backgross.value = 'Labor';
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

    worksheet.mergeCells('L19', 'P19');
    let totalgross = worksheet.getCell('L19');
    totalgross.value = 'Parts';
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
      'Description', 'Count', 'Hours', 'Sales', 'Gross', 'GP %', 'ELR', 'AVG HRS', 'AVG Sales', 'AVG GP', 'Sales', 'Gross', 'GP %', 'AVG  Sales', 'AVG GP'
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
        d.OpcodeDesc == '' ? '-' : d.OpcodeDesc == null ? '-' : d.OpcodeDesc,
        // d.Brand == '' ? '-' : d.Brand == null ? '-' : d.Brand,
        // d.Labortype == '' ? '-' : d.Labortype == null ? '-' : d.Labortype,
        d.Count == '' ? '-' : d.Count == null ? '-' : d.Count,
        // d.Mileage == '' ? '-' : d.Mileage == null ? '-' : d.Mileage,
        d.Hours == '' ? '-' : d.Hours == null ? '-' : d.Hours,
        d.LaborSale == '' ? '-' : d.LaborSale == null ? '-' : d.LaborSale,
        d.LaborGross == '' ? '-' : d.LaborGross == null ? '-' : d.LaborGross,
        d.LaborGP == '' ? '-' : d.LaborGP == null ? '-' : d.LaborGP + '%',
        d.ELR == '' ? '-' : d.ELR == null ? '-' : d.ELR,
        d.avg_hrs == '' ? '-' : d.avg_hrs == null ? '-' : d.avg_hrs,
        d.avg_LaborSale == '' ? '-' : d.avg_LaborSale == null ? '-' : d.avg_LaborSale,
        d.avg_LaborGP == '' ? '-' : d.avg_LaborGP == null ? '-' : d.avg_LaborGP,
        d.PartsSale == '' ? '-' : d.PartsSale == null ? '-' : d.PartsSale,
        d.PartsGross == '' ? '-' : d.PartsGross == null ? '-' : d.PartsGross,
        d.PartsGP == '' ? '-' : d.PartsGP == null ? '-' : d.PartsGP + '%',
        d.avg_PartsSale == '' ? '-' : d.avg_PartsSale == null ? '-' : d.avg_PartsSale,
        d.avg_PartsGP == '' ? '-' : d.avg_PartsGP == null ? '-' : d.avg_PartsGP,
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
          number == 5 || number == 6 || number == 10 ||
          number == 11 || number == 12 || number == 13 || number == 15 || number == 16) {
          cell.numFmt = '$#,##0';
        }
        if (number == 8) {
          cell.numFmt = '$#,##0.00';

        }
        if (number == 4 && number == 9) {
          cell.numFmt = '0.00';
        }
        if (number != 1) {
          cell.alignment = {
            vertical: 'top',
            horizontal: 'right',
            indent: 1,
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
          const Data2 = worksheet.addRow([
            d1.data2 == '' ? '-' : d1.data2 == null ? '-' : d1.data2,
            d1.OpcodeDesc == '' ? '-' : d1.OpcodeDesc == null ? '-' : d1.OpcodeDesc,
            // d1.Brand == '' ? '-' : d1.Brand == null ? '-' : d1.Brand,
            // d1.Labortype == '' ? '-' : d1.Labortype == null ? '-' : d1.Labortype,
            d1.Count == '' ? '-' : d1.Count == null ? '-' : d1.Count,
            // d1.Mileage == '' ? '-' : d1.Mileage == null ? '-' : d1.Mileage,
            d1.Hours == '' ? '-' : d1.Hours == null ? '-' : d1.Hours,
            d1.LaborSale == '' ? '-' : d1.LaborSale == null ? '-' : d1.LaborSale,
            d1.LaborGross == '' ? '-' : d1.LaborGross == null ? '-' : d1.LaborGross,
            d1.LaborGP == '' ? '-' : d1.LaborGP == null ? '-' : d1.LaborGP + '%',
            d1.ELR == '' ? '-' : d1.ELR == null ? '-' : d1.ELR,
            d1.avg_hrs == '' ? '-' : d1.avg_hrs == null ? '-' : d1.avg_hrs,
            d1.avg_LaborSale == '' ? '-' : d1.avg_LaborSale == null ? '-' : d1.avg_LaborSale,
            d1.avg_LaborGP == '' ? '-' : d1.avg_LaborGP == null ? '-' : d1.avg_LaborGP,
            d1.PartsSale == '' ? '-' : d1.PartsSale == null ? '-' : d1.PartsSale,
            d1.PartsGross == '' ? '-' : d1.PartsGross == null ? '-' : d1.PartsGross,
            d1.PartsGP == '' ? '-' : d1.PartsGP == null ? '-' : d1.PartsGP + '%',
            d1.avg_PartsSale == '' ? '-' : d1.avg_PartsSale == null ? '-' : d1.avg_PartsSale,
            d1.avg_PartsGP == '' ? '-' : d1.avg_PartsGP == null ? '-' : d1.avg_PartsGP,
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
              number == 5 || number == 6 || number == 10 ||
              number == 11 || number == 12 || number == 13 || number == 15 || number == 16) {
              cell.numFmt = '$#,##0';
            }
            if (number == 8) {
              cell.numFmt = '$#,##0.00';

            }
            if (number == 4 && number == 9) {
              cell.numFmt = '0.00';
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
                d2.OpcodeDesc == '' ? '-' : d2.OpcodeDesc == null ? '-' : d2.OpcodeDesc,
                // d2.Brand == '' ? '-' : d2.Brand == null ? '-' : d2.Brand,
                // d2.Labortype == '' ? '-' : d2.Labortype == null ? '-' : d2.Labortype,
                d2.Count == '' ? '-' : d2.Count == null ? '-' : d2.Count,
                // d2.Mileage == '' ? '-' : d2.Mileage == null ? '-' : d2.Mileage,
                d2.Hours == '' ? '-' : d2.Hours == null ? '-' : d2.Hours,
                d2.LaborSale == '' ? '-' : d2.LaborSale == null ? '-' : d2.LaborSale,
                d2.LaborGross == '' ? '-' : d2.LaborGross == null ? '-' : d2.LaborGross,
                d2.LaborGP == '' ? '-' : d2.LaborGP == null ? '-' : d2.LaborGP + '%',
                d2.ELR == '' ? '-' : d2.ELR == null ? '-' : d2.ELR,
                d2.avg_hrs == '' ? '-' : d2.avg_hrs == null ? '-' : d2.avg_hrs,
                d2.avg_LaborSale == '' ? '-' : d2.avg_LaborSale == null ? '-' : d2.avg_LaborSale,
                d2.avg_LaborGP == '' ? '-' : d2.avg_LaborGP == null ? '-' : d2.avg_LaborGP,
                d2.PartsSale == '' ? '-' : d2.PartsSale == null ? '-' : d2.PartsSale,
                d2.PartsGross == '' ? '-' : d2.PartsGross == null ? '-' : d2.PartsGross,
                d2.PartsGP == '' ? '-' : d2.PartsGP == null ? '-' : d2.PartsGP + '%',
                d2.avg_PartsSale == '' ? '-' : d2.avg_PartsSale == null ? '-' : d2.avg_PartsSale,
                d2.avg_PartsGP == '' ? '-' : d2.avg_PartsGP == null ? '-' : d2.avg_PartsGP,
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
                if (
                  number == 5 || number == 6 || number == 10 ||
                  number == 11 || number == 12 || number == 13 || number == 15 || number == 16) {
                  cell.numFmt = '$#,##0';
                }
                if (number == 8) {
                  cell.numFmt = '$#,##0.00';

                }
                if (number == 4 && number == 9) {
                  cell.numFmt = '0.00';
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
    worksheet.getColumn(2).width = 40;
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
    workbook.xlsx.writeBuffer().then((data: any) => {
      const blob = new Blob([data], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      this.shared.exportToExcel(workbook, 'OP Code Tracker V2_' + DATE_EXTENSION);
    });
    // })
  }


}
