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

@Component({
  selector: 'app-dashboard',
  imports: [SharedModule, BsDatepickerModule, DateRangePicker, Stores],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
})
export class Dashboard {
  ServiceData: any = []
  IndividualServiceGross: any = [];
  TotalServiceGross: any = [];
  TotalReport: any = 'T'
  NoData: boolean = false;
  responcestatus: any = '';
  LaborTypeVal: any = ''
  LaborState: any = 'S';
  labourType: any = 'N';


  reportOpenSub!: Subscription;
  reportGetting!: Subscription;
  Pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;

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
    { "ARG_ID": 49, "ARG_LABEL": "Store", "ARG_SEQ": 1, "state": true, "columnname": "Store_Name", "Active": "Y" },
    { "ARG_ID": 50, "ARG_LABEL": "Tech Number", "ARG_SEQ": 2, "state": false, "columnname": "ASD_techno", "Active": "Y" },
    { "ARG_ID": 51, "ARG_LABEL": "Tech Name", "ARG_SEQ": 3, "state": true, "columnname": "asd_techname", "Active": "Y" },
    { "ARG_ID": 52, "ARG_LABEL": "Pay Type", "ARG_SEQ": 3, "state": false, "columnname": "ASD_PAYTYPE", "Active": "Y" }
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
    this.selectedDataGrouping.push(this.dataGrouping[2])
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
    this.initializeDates('MTD')
    this.shared.setTitle(this.comm.titleName + '-Service Tech Efficiency');

    this.setHeaderData()
    this.getlabourTypesData('FL', 'N');

  }
  ngOnInit(): void {

  }


  setHeaderData() {
    const data = {
      title: 'Service Tech Efficiency',
      stores: this.storeIds,
      fromdate: this.FromDate,
      todate: this.ToDate,
      groups: this.groupId,
      labortype: this.LaborTypeVal.toString(),
      laborstate: this.LaborState,
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
      if (this.LaborTypeVal == '') {
        const obj = {
          StoreID: this.storeIds.toString(),
          type: this.LaborState
        };
        this.shared.api.postmethod(this.comm.routeEndpoint + 'GetLaborTypesTechEfficiency', obj).subscribe((res) => {
          this.LaborTypeVal = res.response.map(function (a: any) {
            return a.ASD_labortype;
          });
          this.GetData();
          this.GetTotalData();
        })
      } else {
        this.GetData();
        this.GetTotalData();
      }
    } else {
      this.NoData = true
    }
  }
  GetData() {
    this.IndividualServiceGross = [];
    const obj = {
      startdate: this.FromDate,
      enddate: this.ToDate,
      StoreID: this.storeIds == 0 ? '' : this.storeIds.toString(),
      TechNumber: "",
      TechName: "",
      var1: this.selectedDataGrouping.length >= 1 ? this.selectedDataGrouping[0]?.columnname : '',
      var2: this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '',
      var3: this.selectedDataGrouping.length == 3 ? this.selectedDataGrouping[2]?.columnname : '',
      type: "D",
      labortypes: this.LaborTypeVal.toString()
    };
    const curl = environment.apiUrl + this.comm.routeEndpoint + 'GetServiceTechEfficiencyNew';
    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetServiceTechEfficiencyNew', obj).subscribe(
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
      StoreID: this.storeIds == 0 ? '' : this.storeIds,
      TechNumber: "",
      TechName: "",
      var1: this.selectedDataGrouping[0].columnname == 'Store_Name' ? '' : this.selectedDataGrouping[0].columnname,
      var2: this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '',
      var3: this.selectedDataGrouping.length == 3 ? this.selectedDataGrouping[2]?.columnname : '',
      type: "T",
      labortypes: this.LaborTypeVal.toString()
    };
    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetServiceTechEfficiencyNew', obj).subscribe(
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
        //console.log(this.ServiceData);
      } else {
        this.IndividualServiceGross.unshift(this.TotalServiceGross[0]);
        this.ServiceData = this.IndividualServiceGross;
        //console.log(this.ServiceData);
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
  expandorcollapse(ind: any, e: any, ref: any, Item: any, parentData: any, tmp: any) {
    let id = (e.target as Element).id;
    if (id == 'D_' + ind) {
      if ((this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '') == '') {
        this.openDetails(Item, parentData, '1', tmp);
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
        this.openDetails(Item, parentData, '2', tmp);
      } else {
        if (ref == '-') {
          Item.data2sign = '+';
          Item.SubData = Item.SubData.map((obj: any, i: any) => ({ ...obj, data3sign: '+' }))
        }
        if (ref == '+') {
          Item.data2sign = '-';
          Item.Dealer = '-';
        }
      }
    }
    if (id == 'DNS_' + ind) {
      if (ref == '-') {
        Item.data3sign = '+';
      }
      if (ref == '+') {
        Item.data3sign = '-';
        Item.Dealer = '-';
        Item.data2sign = '-'
      }
    }
    // }
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

  Salesdetails: any = []
  openDetails(Item: any, ParentItem: any, ref: any, tmp: any) {
    const DetailsSalesPeron = this.shared.ngbmodal.open(
      tmp,
      {
        size: 'xxl',
        backdrop: 'static',
      }
    );
    this.Salesdetails = [
      {
        StartDate: this.FromDate,
        EndDate: this.ToDate,
        var1: this.selectedDataGrouping[0]?.columnname,
        var2: this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '',
        var3: (this.selectedDataGrouping.length == 3 ? this.selectedDataGrouping[2]?.columnname : ''),
        StoreID: this.selectedDataGrouping[0]?.columnname == 'Store_Name' ? (ParentItem.StoreID) : (this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '') == 'Store_Name' ? (Item.StoreID) : (this.selectedDataGrouping.length == 3 ? this.selectedDataGrouping[2]?.columnname : '') == 'Store_Name' ? (Item.StoreID) : '',
        TechNumber: this.selectedDataGrouping[0]?.columnname == 'ASD_techno' ? (ParentItem.data1) : (this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '') == 'ASD_techno' ? (Item.data2) : (this.selectedDataGrouping.length == 3 ? this.selectedDataGrouping[2]?.columnname : '') == 'ASD_techno' ? (Item.data3) : '',
        TechName: this.selectedDataGrouping[0]?.columnname == 'asd_techname' ? (ParentItem.data1) : (this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '') == 'asd_techname' ? (Item.data2) : (this.selectedDataGrouping.length == 3 ? this.selectedDataGrouping[2]?.columnname : '') == 'asd_techname' ? (Item.data3) : '',
        paytype: this.selectedDataGrouping[0]?.columnname == 'ASD_PAYTYPE' ? (ParentItem.data1) : (this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '') == 'ASD_PAYTYPE' ? (Item.data2) : (this.selectedDataGrouping.length == 3 ? this.selectedDataGrouping[2]?.columnname : '') == 'ASD_PAYTYPE' ? (Item.data3) : '',
        StoreName: this.selectedDataGrouping[0]?.columnname == 'Store_Name' ? (ParentItem.data1) : (this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '') == 'Store_Name' ? (Item.data2) : (this.selectedDataGrouping.length == 3 ? this.selectedDataGrouping[2]?.columnname : '') == 'Store_Name' ? (Item.data3) : '',
        laborvalues: this.LaborTypeVal.toString()
      },
    ];
    this.getDetails()
  }

  NoDetailsData!: boolean;
  spinnerLoader: boolean = true;
  spinnerLoadersec: boolean = false;
  pageNumber: any = 0;

  viewRO(roData: any) {
    // const modalRef = this.ngbmodel.open(RepairOrderComponent, { size: 'md', windowClass: 'connectedmodal' });
    // modalRef.componentInstance.data = {  ro: roData.RO, storeid: roData.StoreID, vin:roData.vin, vehicleid:roData.vehicleid,custno: roData?.customernumber  }; // Pass data to the modal component    
    // modalRef.result.then((result) => {
    //   console.log(result); // Handle modal close result
    // }, (reason) => {
    //   console.log(`Dismissed: ${reason}`); // Handle dismiss reason
    // });
  }
  details: any = [];
  ServiceDataDetails: any = []
  getDetails() {
    this.ServiceDataDetails = []
    const obj = {
      startdate: this.Salesdetails[0].StartDate,
      enddate: this.Salesdetails[0].EndDate,
      StoreID: this.Salesdetails[0].StoreID ? this.Salesdetails[0].StoreID : '',
      TechNumber: this.Salesdetails[0].TechNumber,
      TechName: this.Salesdetails[0].TechName,
      paytype: this.Salesdetails[0].paytype,
      var1: this.Salesdetails[0].var1,
      var2: this.Salesdetails[0].var2,
      var3: this.Salesdetails[0].var3,
      type: "D",
      labortypes: this.Salesdetails[0].laborvalues
    }
    this.spinnerLoader = true;
    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetServiceTechEfficiencyDetailsV1', obj).subscribe(
      (res) => {
        //console.log(res);
        if (res.status == 200) {
          this.spinnerLoader = false
          if (res.response != undefined) {
            if (res.response.length > 0) {
              this.ServiceDataDetails = res.response
              this.ServiceDataDetails.some(function (x: any) {
                if (x.Data2 != undefined) {
                  x.Data2 = JSON.parse(x.Data2);
                  x.Dealer = '+'
                }
              });
              this.NoData = false;
            } else {
              this.NoData = true;
            }

          }
        }

        //console.log(this.ServiceData);

      })
  }

  closePopup() {
    this.shared.ngbmodal.dismissAll()
  }

  async openServiceModal(roNumber: any, vin: any, storeid: any, vehicleid: any, source: any, custno: any) {
    const module = await import('../../../../Layout/cdpdataview/repair/repair-module');
    const component = module.Repair;
    const modalRef = this.shared.ngbmodal.open(component, { size: 'xl', windowClass: 'connectedmodal' });
    modalRef.componentInstance.data = { ro: roNumber, vin: vin, storeid: storeid, vehicleid: vehicleid, source: source, custno: custno }; // Pass data to the modal component
    modalRef.result.then((result) => {
      console.log(result); // Handle modal close result
    }, (reason) => {
      console.log(`Dismissed: ${reason}`); // Handle dismiss reason
    });
  }
  currentElement!: string;

  @ViewChild('scrollOne') scrollOne!: ElementRef;
  @ViewChild('scrollTwo') scrollTwo!: ElementRef;

  updateVerticalScroll(event: any): void {
    if (this.currentElement === 'scrollTwo') {
      this.scrollOne.nativeElement.scrollTop = event.target.scrollTop;
      if (
        event.target.scrollTop + event.target.clientHeight >=
        event.target.scrollHeight - 2
      ) {
        // alert("reached at bottom");
        if (this.pageNumber == 0) {
          this.spinnerLoadersec = true;
          this.pageNumber++;
          this.getDetails();
        } else {
          if (this.details.length > 0) {
            this.spinnerLoadersec = true;
            this.pageNumber++;
            this.getDetails();
          }
        }
      }
    } else if (this.currentElement === 'scrollOne') {
      this.scrollTwo.nativeElement.scrollTop = event.target.scrollTop;
      if (
        event.target.scrollTop + event.target.clientHeight >=
        event.target.scrollHeight - 2
      ) {
        if (this.pageNumber == 0) {
          this.spinnerLoadersec = true;
          this.pageNumber++;
          this.getDetails();
        } else {
          if (this.details.length > 0) {
            this.spinnerLoadersec = true;
            this.pageNumber++;
            this.getDetails();
          }
        }
      }
    }
  }

  updateCurrentElement(element: 'scrollOne' | 'scrollTwo') {
    this.currentElement = element;
  }


  ngAfterViewInit() {
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.comm.pageName == 'Service Tech Efficiency') {
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
        if (res.obj.title == 'Service Tech Efficiency') {
          if (res.obj.state == true) {
            this.exportToExcel();
          }
        }
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res) => {
      if (this.print != undefined) {
        if (res.obj.title == 'Service Tech Efficiency') {
          if (res.obj.statePrint == true) {
            //  this.GetPrintData();
          }
        }
      }
    });


    this.Pdf = this.shared.api.getExportToPDFAllReports().subscribe((res) => {
      if (this.Pdf != undefined) {
        if (res.obj.title == 'Service Tech Efficiency') {
          if (res.obj.statePDF == true) {
            //  this.generatePDF();
          }
        }
      }
    });
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res) => {
      if (this.email != undefined) {
        if (res.obj.title == 'Service Tech Efficiency') {
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
  labortypes: any = []
  StoresData(data: any) {
    this.storeIds = data.storeids;
    this.groupId = data.groupId;
    this.storename = data.storename;
    this.groupName = data.groupName;
    this.storecount = data.storecount;
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


    if (block == 'LT') {
      this.labourType = e
      this.getlabourTypesData('FR', e)
    }


  }

  getlabourTypesData(block: any, type: any) {
    if (this.storeIds != '') {
      this.spinnerLoaderlabor = true;
      this.responcestatus = '';

      // this.shared.spinner.show();    
      const obj = {
        StoreID: this.storeIds.toString(),
        type: type == 'N' ? 'A' : type
      };
      this.shared.api.postmethod(this.comm.routeEndpoint + 'GetLaborTypesTechEfficiency', obj).subscribe((res) => {
        this.spinnerLoaderlabor = false;
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

  spinnerLoaderlabor: boolean = false;
  activePopover: number = -1;

  togglePopover(popoverIndex: number) {
    this.activePopover = this.activePopover === popoverIndex ? -1 : popoverIndex;
  }

  viewreport() {
    this.activePopover = -1
    if (this.selectedDataGrouping.length == 0 || this.LaborTypeVal.length == 0 || this.storeIds.length == 0) {
      if (this.selectedDataGrouping.length == 0) {
        this.toast.show('Please select atleast one Value from Grouping', 'warning', 'Warning');
      }
      if (this.storeIds.length == 0) {
        this.toast.show('Please select atleast one Store', 'warning', 'Warning');
      }
      if (this.LaborTypeVal.length == 0) {
        this.toast.show('Please select any labor type', 'warning', 'Warning');
      }
    } else {

      this.setHeaderData()
      this.getServiceData()
    }
  }

  ExcelStoreNames: any = [];
  exportToExcel() {
    let storeNames: any = [];
    let store = this.storeIds
    storeNames = this.comm.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores.filter((item: any) =>
      store.some((cat: any) => cat === item.ID.toString())
    );
    if (store.length == this.comm.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores.length) {
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
    const worksheet = workbook.addWorksheet('Service Tech Efficiency');
    worksheet.views = [
      {
        showGridLines: false,
      },
    ];
    worksheet.addRow('');
    const titleRow = worksheet.addRow(['Service Tech Efficiency']);
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
    groupings.value = this.selectedDataGrouping[0]?.ARG_LABEL + ',' + this.selectedDataGrouping[1]?.ARG_LABEL + ',' + this.selectedDataGrouping[2]?.ARG_LABEL;
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
    groups.value = this.comm.groupsandstores.filter((val: any) => val.sg_id == this.groupId.toString())[0].sg_name;
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
    let units = worksheet.getCell('D16');
    units.value = '';
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

    worksheet.mergeCells('E16', 'G16');
    let backgross = worksheet.getCell('E16');
    backgross.value = 'Productivity';
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
    worksheet.mergeCells('H16', 'J16');
    let totalgross = worksheet.getCell('H16');
    totalgross.value = 'Efficiency';
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
    //   let dateYear = worksheet.getCell('A16');
    let Productivity = worksheet.getCell('K16');
    Productivity.value = '';
    Productivity.alignment = { vertical: 'middle', horizontal: 'center' };
    Productivity.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
      color: { argb: 'FFFFFF' },
    };
    Productivity.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '2a91f0' },
      bgColor: { argb: 'FF0000FF' },
    };
    Productivity.border = { right: { style: 'thin' } };

    let Headings = [
      FromDate + ' - ' + ToDate + ', ' + PresentYear,
      'RO Count',
      'Line Count	',
      'Unapplied Time Cost',
      // 'Bill Amt',
      // 'Job Cost',
      'Flag Hrs',
      'Time Card Hrs',
      ' % Flag/Time Card',
      'Flag Hrs',
      'Actual Hrs on Jobs',
      '% Flag /Actual',
      // 'Time Card Hours',
      // '% Time Card / Actual',
      // 'Idle Hours',
      // '%'
      'ELR'
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
      let obj = [
        (d.data1 == undefined || d.data1 == '' || d.data1 == ' ' ? '--' : (d.data1 == 'C' ?
          'Customer Pay' : (d.data1 == 'T' ? 'Warranty' : (d.data1 ==
            'I' ? 'Internal' : (d.data1 == 'R' ? 'Retail' : (d.data1 ==
              'W' ? 'Warranty' : ((d.data1)))))))),
        d.RoCount == '' ? '-' : d.RoCount == null ? '-' : d.RoCount,
        d.Linecount == '' ? '-' : d.Linecount == null ? '-' : d.Linecount,
        d.unappliedcost == '' ? '-' : d.unappliedcost == null ? '-' : d.unappliedcost,
        // d.Bill == '' ? '-' : d.Bill == null ? '-' : d.Bill,
        // d.Jobcost == '' ? '-' : d.Jobcost == null ? '-' : d.Jobcost,
        d.flaghours == '' ? '-' : d.flaghours == null ? '-' : d.flaghours,
        d.ASD_timecardhours == '' ? '-' : d.ASD_timecardhours == null ? '-' : d.ASD_timecardhours,
        d.proficiency == '' ? '-' : d.proficiency == null ? '-' : d.proficiency + '%',
        d.flaghours == '' ? '-' : d.flaghours == null ? '-' : d.flaghours,
        d.Actualhours == '' ? '-' : d.Actualhours == null ? '-' : d.Actualhours,
        d.efficiency == '' ? '-' : d.efficiency == null ? '-' : d.efficiency + '%',
        // d.ASD_timecardhours == '' ? '-' : d.ASD_timecardhours == null ? '-' : d.ASD_timecardhours,
        // d.productivity == '' ? '-' : d.productivity == null ? '-' : d.productivity + '%',
        // '-', '-'
        d.ELR == '' ? '-' : d.ELR == null ? '-' : d.ELR,

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
        if (number == 2 || number == 3 || number == 4) {
          cell.numFmt = '#,##0';
        } if (number == 5 || number == 6) {
          cell.numFmt = '$#,##0';
        } if (number >= 7) {
          cell.numFmt = '#,##0.00';
        } if (number != 1) {
          cell.alignment = { vertical: 'middle', horizontal: 'center', indent: 1 };
        }
        if (number > 1 && obj[number] != undefined) {
          if (obj[number] < 0) {
            Data1.getCell(number + 1).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } };
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
      if (d.Data2 != undefined) {
        for (const dt1 of d.Data2) {
          let obj = [
            (dt1.data2 == undefined || dt1.data2 == '' || dt1.data2 == ' ' ? '--' : (dt1.data2 == 'C' ?
              'Customer Pay' : (dt1.data2 == 'T' ? 'Warranty' : (dt1.data2 ==
                'I' ? 'Internal' : (dt1.data2 == 'R' ? 'Retail' : (dt1.data2 ==
                  'W' ? 'Warranty' : ((dt1.data2)))))))),
            dt1.RoCount == '' ? '-' : dt1.RoCount == null ? '-' : dt1.RoCount,
            dt1.Linecount == '' ? '-' : dt1.Linecount == null ? '-' : dt1.Linecount,
            dt1.unappliedcost == '' ? '-' : dt1.unappliedcost == null ? '-' : dt1.unappliedcost,
            // dt1.Bill == '' ? '-' : dt1.Bill == null ? '-' : dt1.Bill,
            // dt1.Jobcost == '' ? '-' : dt1.Jobcost == null ? '-' : dt1.Jobcost,
            dt1.flaghours == '' ? '-' : dt1.flaghours == null ? '-' : dt1.flaghours,
            dt1.ASD_timecardhours == '' ? '-' : dt1.ASD_timecardhours == null ? '-' : dt1.ASD_timecardhours,
            dt1.proficiency == '' ? '-' : dt1.proficiency == null ? '-' : dt1.proficiency + '%',
            dt1.flaghours == '' ? '-' : dt1.flaghours == null ? '-' : dt1.flaghours,
            dt1.Actualhours == '' ? '-' : dt1.Actualhours == null ? '-' : dt1.Actualhours,
            dt1.efficiency == '' ? '-' : dt1.efficiency == null ? '-' : dt1.efficiency + '%',
            // dt1.ASD_timecardhours == '' ? '-' : dt1.ASD_timecardhours == null ? '-' : dt1.ASD_timecardhours,
            // dt1.productivity == '' ? '-' : dt1.productivity == null ? '-' : dt1.productivity + '%',
            // '-', '-'
            dt1.ELR == '' ? '-' : dt1.ELR == null ? '-' : dt1.ELR,

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
            if (number == 2 || number == 3 || number == 4) {
              cell.numFmt = '#,##0';
            } if (number == 5 || number == 6) {
              cell.numFmt = '$#,##0';
            } if (number >= 7) {
              cell.numFmt = '#,##0.00';
            } if (number != 1) {
              cell.alignment = { vertical: 'middle', horizontal: 'center', indent: 1 };
            }
            if (number > 1 && obj[number] != undefined) {
              if (obj[number] < 0) {
                Data1.getCell(number + 1).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } };
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
          if (dt1.SubData != undefined) {
            for (const d2 of dt1.SubData) {
              let obj = [
                (d2.data3 == undefined || d2.data3 == '' || d2.data3 == ' ' ? '--' : (d2.data3 == 'C' ?
                  'Customer Pay' : (d2.data3 == 'T' ? 'Warranty' : (d2.data3 ==
                    'I' ? 'Internal' : (d2.data3 == 'R' ? 'Retail' : (d2.data3 ==
                      'W' ? 'Warranty' : ((d2.data3)))))))),
                d2.RoCount == '' ? '-' : d2.RoCount == null ? '-' : d2.RoCount,
                d2.Linecount == '' ? '-' : d2.Linecount == null ? '-' : d2.Linecount,
                d2.unappliedcost == '' ? '-' : d2.unappliedcost == null ? '-' : d2.unappliedcost,
                // d2.Bill == '' ? '-' : d2.Bill == null ? '-' : d2.Bill,
                // d2.Jobcost == '' ? '-' : d2.Jobcost == null ? '-' : d2.Jobcost,
                d2.flaghours == '' ? '-' : d2.flaghours == null ? '-' : d2.flaghours,
                d2.ASD_timecardhours == '' ? '-' : d2.ASD_timecardhours == null ? '-' : d2.ASD_timecardhours,
                d2.proficiency == '' ? '-' : d2.proficiency == null ? '-' : d2.proficiency + '%',
                d2.flaghours == '' ? '-' : d2.flaghours == null ? '-' : d2.flaghours,
                d2.Actualhours == '' ? '-' : d2.Actualhours == null ? '-' : d2.Actualhours,
                d2.efficiency == '' ? '-' : d2.efficiency == null ? '-' : d2.efficiency + '%',
                // d2.ASD_timecardhours == '' ? '-' : d2.ASD_timecardhours == null ? '-' : d2.ASD_timecardhours,
                // d2.productivity == '' ? '-' : d2.productivity == null ? '-' : d2.productivity + '%',
                // '-', '-'

                d2.ELR == '' ? '-' : d2.ELR == null ? '-' : d2.ELR,

              ]
              const Data3 = worksheet.addRow(obj);
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
                if (number == 2 || number == 3 || number == 4) {
                  cell.numFmt = '#,##0';
                } if (number == 5 || number == 6) {
                  cell.numFmt = '$#,##0';
                } if (number >= 7) {
                  cell.numFmt = '#,##0.00';
                } if (number != 1) {
                  cell.alignment = { vertical: 'middle', horizontal: 'center', indent: 1 };
                }
                if (number > 1 && obj[number] != undefined) {
                  if (obj[number] < 0) {
                    Data1.getCell(number + 1).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } };
                  }
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
      if (d.data1 === 'REPORTS TOTAL') {
        Data1.eachCell((cell) => {
          cell.font = { name: 'Arial', family: 4, size: 9, bold: true };
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
    worksheet.getColumn(19).width = 15;
    worksheet.getColumn(20).width = 15;
    worksheet.getColumn(21).width = 15;
    worksheet.addRow([]);

    this.shared.exportToExcel(workbook, 'Service Tech Efficiency_' + DATE_EXTENSION);

  }


}
