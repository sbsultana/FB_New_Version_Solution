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

  DupFromDate: any = '';
  DupToDate: any = ''


  SalesData: any = [];
  IndividualTires: any = [];
  TotalTires: any = [];
  IndividualSalesGross: any = [];
  TotalSalesGross: any = [];
  TotalSortPosition: any = 'B';
  NoData: any = '';

  TotalReport: any = ['B'];
  responcestatus: string = '';

  reportOpenSub!: Subscription;
  reportGetting!: Subscription;
  Pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;
  // Report Popup

  scrollLoader: boolean = false;
  dataGrouping: any = [
    { "ARG_ID": 40, "ARG_LABEL": "Store", "ARG_SEQ": 0, 'Column_Name': 'DealerName' },
    { "ARG_ID": 27, "ARG_LABEL": "Advisor Name", "ARG_SEQ": 2, 'Column_Name': 'Ap_sa_name' },
    { "ARG_ID": 28, "ARG_LABEL": "Tech Name", "ARG_SEQ": 3, 'Column_Name': 'TechName' },
    { "ARG_ID": 26, "ARG_LABEL": "Counter Person", "ARG_SEQ": 1, 'Column_Name': 'AP_CounterPerson' },
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
    this.shared.setTitle(this.comm.titleName + '-Tire Group Sales Report');
    this.grouping.push(this.dataGrouping[0]);


    this.initializeDates('MTD')

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

    this.setHeaderData()
    this.getSalesData();
  }
  ngOnInit(): void {
  }


  datetype() {
    if (this.DateType == 'PM') {
      return 'SP';
    }
    else if (this.DateType == 'C') {
      return 'C'
    }
    return this.DateType;
  }

  initializeDates(type: any) {
    let dates: any = this.setdates.setDates(type)
    this.FromDate = dates[0];
    this.ToDate = dates[1];
    localStorage.setItem('time', type);
    this.DateType = type
    this.setDates(this.DateType)

  }
  getSalesData() {
    if (this.storeIds != undefined && this.storeIds.length > 0) {
      this.shared.spinner.show();
      this.GetData();
      this.GetTotalsData();
    } else {
      // this.NoData = true;
      this.shared.spinner.hide();
    }
  }
  GetData() {

    console.log(this.grouping);

    this.IndividualSalesGross = [];
    this.DupFromDate = this.FromDate;
    this.DupToDate = this.ToDate
    const obj = {
      startdate: this.FromDate,
      enddate: this.ToDate,
      Store: this.storeIds.toString(),
      Labortype: this.Department.indexOf('Service') >= 0 ? this.servicetype.toString() : '',
      Saletype: this.Department.indexOf('Parts') >= 0 ? this.saletype.toString() : '',
      DepartmentS: this.Department.indexOf('Service') >= 0 ? 'S' : '',
      DepartmentP: this.Department.indexOf('Parts') >= 0 ? 'P' : '',
      "var1": this.grouping[0].Column_Name,
      "var2": this.grouping.length > 1 ? this.grouping[1].Column_Name : '',
      "var3": "",
      Type: 'D'
    };
    const curl = environment.apiUrl + 'fredbeans/GetTireGroupSalesReportNewV1';
    this.shared.api
      .postmethod(this.comm.routeEndpoint + 'GetTireGroupSalesReportNewV1', obj)
      .subscribe(
        (res) => {
          const currentTitle = document.title;
          this.shared.api.logSaving(curl, {}, '', res.message, currentTitle);
          this.IndividualTires = [];
          this.SalesData = [];
          if (res && res.response && res.response.length > 0) {
            this.responcestatus = this.responcestatus + 'I';
            this.IndividualTires = res.response.map((v: any) => ({ ...v, Dealer: '+', }));
            this.IndividualTires.some(function (x: any) {
              if (x.Data2 != undefined && x.Data2 != '' && x.Data2 != null) {
                x.Data2 = JSON.parse(x.Data2);
                x.Data2 = x.Data2.map((v: any) => ({ ...v, SubData: [], data2sign: '-', }));
              }
            });
            this.combineIndividualandTotal();
            this.shared.spinner.hide();
            this.NoData = '';
          } else {
            this.shared.spinner.hide();
            this.NoData = 'No Data Found!!';
          }
        },
        (error) => {
          this.toast.show('502 Bad Gate Way Error', 'danger','Error');
          this.shared.spinner.hide();
          this.NoData = 'No Data Found';
        }
      );
  }
  GetTotalsData() {
    this.IndividualSalesGross = [];
    this.DupFromDate = this.FromDate;
    this.DupToDate = this.ToDate
    const obj = {
      startdate: this.FromDate,
      enddate: this.ToDate,
      Store: this.storeIds.toString(),
      Labortype: this.Department.indexOf('Service') >= 0 ? this.servicetype.toString() : '',
      Saletype: this.Department.indexOf('Parts') >= 0 ? this.saletype.toString() : '',
      DepartmentS: this.Department.indexOf('Service') >= 0 ? 'S' : '',
      DepartmentP: this.Department.indexOf('Parts') >= 0 ? 'P' : '',
      "var1": this.grouping[0].Column_Name,
      "var2": this.grouping.length > 1 ? this.grouping[1].Column_Name : '',
      "var3": "",
      Type: 'T'
    };
    const curl = environment.apiUrl + 'fredbeans/GetTireGroupSalesReportNewV1';
    this.shared.api
      .postmethod(this.comm.routeEndpoint + 'GetTireGroupSalesReportNewV1', obj)
      .subscribe(
        (res) => {
          const currentTitle = document.title;
          this.shared.api.logSaving(curl, {}, '', res.message, currentTitle);
          this.TotalTires = [];
          if (res && res.response && res.response.length > 0) {
            this.responcestatus = this.responcestatus + 'T';
            this.TotalTires = res.response;
            this.combineIndividualandTotal();
            // this.shared.spinner.hide();
            this.NoData = '';
          } else {
            this.shared.spinner.hide();
            this.NoData = 'No Data Found!!';
          }
        },
        (error) => {
          this.toast.show('502 Bad Gate Way Error', 'danger','Error');
          this.shared.spinner.hide();
          this.NoData = 'No Data Found';
        }
      );
  }
  combineIndividualandTotal() {
    this.SalesData = this.IndividualTires;
    // this.shared.spinner.hide();
    console.log(this.TotalReport.toString(), '..........');
    if (this.responcestatus == 'IT' || this.responcestatus == 'TI') {
      if (this.TotalReport.toString() == 'B') {
        this.IndividualTires.push(this.TotalTires[0]);
        this.SalesData = this.IndividualTires;
        this.shared.spinner.hide();
        console.log(this.SalesData);
      } else {
        console.log(this.IndividualTires, this.TotalTires, '...........');
        this.IndividualTires.unshift(this.TotalTires[0]);
        this.SalesData = this.IndividualTires;
        this.shared.spinner.hide();
        console.log(this.SalesData);
      }
    } else if (this.responcestatus == 'T') {
      this.SalesData = this.TotalTires;
      this.shared.spinner.hide();
    } else if (this.responcestatus == 'I') {
      this.SalesData = this.IndividualTires;
      this.shared.spinner.hide();
    } else {
      this.NoData = 'No Data Found!!';
    }
    console.log(this.SalesData, 'Sales Data');
  }
  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    } else if (value < 0) {
      return false;
    }
    return true;
  }
  isDesc: boolean = false;
  column: string = 'CategoryName';
  sort(property: any) {
    this.isDesc = !this.isDesc; //change the direction
    this.column = property;
    let direction = this.isDesc ? 1 : -1;
    if (direction == -1) {
      this.TotalSortPosition = 'T';
    } else {
      this.TotalSortPosition = 'B';
    }
    this.SalesData.sort(function (a: any, b: any) {
      if (a[property] < b[property]) {
        return -1 * direction;
      } else if (a[property] > b[property]) {
        return 1 * direction;
      } else {
        return 0;
      }
    });
  }
  expandorcollapse(ind: any, e: any, ref: any, Item: any, parentData: any) {
    let id = (e.target as Element).id;
    if (id == 'D_' + ind) {
      if (this.grouping.length == 1) {
        this.reconaccountpopup(parentData, Item, '');
      } else {
        if (ref == '-') {
          Item.Dealer = '+';
        }
        if (ref == '+') {
          Item.Dealer = '-';
        }
      }
    }
  }
  popupReference!: NgbModalRef;
  popupvalues: any = { Acctno: '', AcctDesc: '', Dept: '', Store: '', Storename: '', subtype: '' }
  ReconDetails: any = []
  Details: any = []
  ReconDetailsNoData: any = '';
  pageNumber: any = 0;
  data: any = {}
  parent: any = {}
  reconaccountpopup(tmp: any, data: any, parent: any) {
    this.data = data;
    this.parent = parent;
    // console.log(data, parent);
    // this.popupvalues.FIN = FIN;
    // this.popupvalues.Acctno = data.accountnumber;
    // this.popupvalues.AcctDesc = data.accountdescription;
    this.popupvalues.Store = data.Store_Name ? data.Store_Name : parent.Store_Name;
    // this.popupvalues.Storename = StoreName;
    this.popupvalues.subtype = data.data2 ? data.data2 : '';
    this.ReconDetailsNoData = '';
    this.ReconDetails = []
    this.popupReference = this.shared.ngbmodal.open(tmp, { size: 'xl', backdrop: 'static', keyboard: true, centered: true, modalDialogClass: 'custom-modal' })
    this.pageNumber = 0
    this.GetReconDetails(data, parent)
  }
  GetReconDetails(data: any, parent: any) {
    console.log(data, parent);

    const obj = {
      startdate: this.FromDate,
      enddate: this.ToDate,
      Store: data.DealerID ? data.DealerID : parent.DealerID,
      Labortype: this.Department.indexOf('Service') >= 0 ? this.servicetype.toString() : '',
      Saletype: this.Department.indexOf('Parts') >= 0 ? this.saletype.toString() : '',
      DepartmentS: this.Department.indexOf('Service') >= 0 ? 'S' : '',
      DepartmentP: this.Department.indexOf('Parts') >= 0 ? 'P' : '',
      var1: this.grouping[0].Column_Name,
      var2: this.grouping.length > 1 ? this.grouping[1].Column_Name : '',
      var3: "",
      var1Value: this.popupvalues.Store,
      var2Value: this.popupvalues.subtype,
      var3Value: "",
      PageNumber: this.pageNumber,
      PageSize: "100"
    }
    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetTireGroupSalesReportDetailsNewV1', obj).subscribe(
      (res) => {
        if (res.response && res.response.length > 0) {
          this.Details = res.response;
          let TiresDetails = [
            ...this.ReconDetails,
            ...this.Details,
          ];
          this.ReconDetails = TiresDetails
          this.ReconDetailsNoData = '';
          this.scrollLoader = false;
        } else {
          this.Details = []
          this.ReconDetailsNoData = 'No Data Found';
          this.scrollLoader = false;

        }
      },
      (error) => {
        this.ReconDetailsNoData = 'No Data Found';
        this.scrollLoader = false;

      })
  }
  closePopup() {
    if (this.popupReference) {
      this.popupReference.close();
    }
  }
  ngAfterViewInit() {
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.comm.pageName == 'Tire Group Sales Report') {
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
        if (res.obj.title == 'Tire Group Sales Report') {
          if (res.obj.state == true) {
            this.exportToExcel();
          }
        }
      }
    });
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res) => {
      if (this.email != undefined) {
        if (res.obj.title == 'Tire Group Sales Report') {
          if (res.obj.stateEmailPdf == true) {
         //   this.sendEmailData(res.obj.Email, res.obj.notes, res.obj.from);
          }
        }
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res) => {
      if (this.print != undefined) {
        if (res.obj.title == 'Tire Group Sales Report') {
          if (res.obj.statePrint == true) {
       //     this.GetPrintData();
          }
        }
      }
    });
    this.Pdf = this.shared.api.getExportToPDFAllReports().subscribe((res) => {
      if (this.Pdf != undefined) {
        console.log('PDF');
        if (res.obj.title == 'Tire Group Sales Report') {
          console.log('PDF Tire Group Sales Report');
          if (res.obj.statePDF == true) {
         //   this.generatePDF();
          }
        }
      }
    });

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

  // @ViewChild('scrollOne') scrollOne!: ElementRef;
  @ViewChild('scrollOne', { static: false }) scrollOne!: ElementRef;
  count = 0
  updateVerticalScroll(event: any): void {
    console.log(event.target.scrollTop + event.target.clientHeight, event.target.scrollHeight - 2, 'Scroll Position', this.Details);
    if (event.target.scrollTop + event.target.clientHeight >= event.target.scrollHeight - 2) {
      if (this.Details.length >= 100) {
        this.scrollLoader = true
        this.pageNumber++;
        this.GetReconDetails(this.data, this.parent);
      }
    }
  }
  getBackgroundColor(value: any, block: any) {
    if (block == 'New' || block == 'Used') {
      if (value > 45 || value < -45) {
        return '#e37474';
      }
      else {
        return '#6fb979';
      }
    } else if (block == 'Service') {
      if (value > 57.5 || value < -57.5) {
        return '#e37474';
      }
      else {
        return '#6fb979';
      }
    } else {
      if (value > 60 || value < -60) {
        return '#e37474';
      }
      else {
        return '#6fb979';
      }
    }
  }
  setHeaderData() {
    const headerdata = {
      title: 'Tire Group Sales Report',
      stores: this.storeIds,
      fromdate: this.FromDate,
      todate: this.ToDate,
      groups: this.groupId,
      Department: this.Department,
      saletype: this.saletype,
      servicetype: this.servicetype,
      path1: this.grouping[0].ARG_LABEL,
      path2: this.grouping.length > 1 ? this.grouping[1].ARG_LABEL : '',
    };
    this.shared.api.SetHeaderData({
      obj: headerdata,
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
  // Report popup Code
 
  Department: any = ['Service', 'Parts'];
  saletype: any = [];
  servicetype: any = ['C', 'T', 'I', 'E']
  saleandservice: any = ['C', 'T', 'I', 'E']
  grouping: any = []
  multipleorsingle(block: any, e: any) {
    if (block == 'Dept') {
      const index = this.Department.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.Department.splice(index, 1);
        if (this.Department.indexOf('Service') >= 0 && this.Department.indexOf('Parts') >= 0) {
          this.saletype = ['R', 'W']
          this.servicetype = ['C', 'T', 'I', 'E']
          this.saleandservice = []
          this.saleandservice = [...this.saletype, ...this.servicetype]
        }
        else if (this.Department.indexOf('Service') >= 0) {
          this.servicetype = ['C', 'T', 'I', 'E']
          this.saletype = []
          this.saleandservice = []
          this.saleandservice = [...this.saletype, ...this.servicetype]
        }
        else if (this.Department.indexOf('Parts') >= 0) {
          this.saletype = ['R', 'W']
          this.servicetype = []
          this.saleandservice = []
          this.saleandservice = [...this.saletype, ...this.servicetype]
        } else {
          this.saletype = []
          this.servicetype = []
          this.saleandservice = []
          this.saleandservice = [...this.saletype, ...this.servicetype]
        }
      } else {
        this.Department.push(e);
        if (this.Department.indexOf('Service') >= 0 && this.Department.indexOf('Parts') >= 0) {
          this.saletype = ['R', 'W']
          this.servicetype = ['C', 'T', 'I', 'E']
          this.saleandservice = []
          this.saleandservice = [...this.saletype, ...this.servicetype]
        }
        else if (this.Department.indexOf('Service') >= 0) {
          this.servicetype = ['C', 'T', 'I', 'E']
          this.saletype = []
          this.saleandservice = []
          this.saleandservice = [...this.saletype, ...this.servicetype]
        }
        else if (this.Department.indexOf('Parts') >= 0) {
          this.saletype = ['R', 'W']
          this.servicetype = []
          this.saleandservice = []
          this.saleandservice = [...this.saletype, ...this.servicetype]
        }
      }
      console.log(this.saleandservice);
    }
    if (block == 'PT') {
      const index = this.saletype.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.saletype.splice(index, 1);
        this.saleandservice = []
        this.saleandservice = [...this.saletype, ...this.servicetype]
      } else {
        this.saletype.push(e);
        this.saleandservice = []
        this.saleandservice = [...this.saletype, ...this.servicetype]
      }
    }
    if (block == 'LT') {
      const index = this.servicetype.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.servicetype.splice(index, 1);
        this.saleandservice = []
        this.saleandservice = [...this.saletype, ...this.servicetype]
      } else {
        this.servicetype.push(e);
        this.saleandservice = []
        this.saleandservice = [...this.saletype, ...this.servicetype]
      }
    }
    if (block == 'TB') {
      this.TotalReport = []
      this.TotalReport.push(e)
    }
    if (block == 'GRP') {

      const index = this.grouping.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.grouping.splice(index, 1);
      } else {
        if (this.grouping.length == 2) {
          this.toast.show('Select up to 2 Filters only to Group your data','warning','Warning');
        } else {
          this.grouping.push(e);

        }
      }


      console.log(this.grouping, 'Groupings');

    }
  }
   activePopover: number = -1;

  togglePopover(popoverIndex: number) {
    this.activePopover = this.activePopover === popoverIndex ? -1 : popoverIndex;
  }
  viewreport() {
    this.activePopover = -1
    if (this.storeIds.length == 0) {
      this.toast.show('Please Select Atleast One Store','warning','Warning');
    } else if (this.Department.length == 0) {
      this.toast.show('Please Select Atleast One Department','warning','Warning');
    }
    else {
      this.responcestatus = ''
      this.setHeaderData();
      this.getSalesData()
    }
  }

  ExcelStoreNames: any = []
  exportToExcel() {
    let storeNames: any = [];
    let store = this.storeIds
    storeNames = this.comm.groupsandstores
      .find((v: any) => v.sg_id == this.groupId)
      ?.Stores.filter((item: any) => store.includes(Number(item.ID)));
    console.log(storeNames);
    if (store.length == this.comm.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores.length) {
      this.ExcelStoreNames = 'All Stores'
    } else {
      this.ExcelStoreNames = storeNames.map(function (a: any) {
        return a.storename;
      });
      console.log(this.ExcelStoreNames);
    }
    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('Tire Group Sales Report');
    worksheet.views = [
      {
        state: 'frozen',
        ySplit: 13, // Number of rows to freeze (2 means the first two rows are frozen)
        topLeftCell: 'A14', // Specify the cell to start freezing from (in this case, the third row)
        showGridLines: true,
      },
    ];
    worksheet.addRow('');
    const titleRow = worksheet.addRow(['Tire Group Sales Report']);
    titleRow.eachCell((cell, number) => {
      cell.alignment = { indent: 1, vertical: 'top', horizontal: 'left' };
    });
    titleRow.font = { name: 'Arial', family: 4, size: 12, bold: true };
    titleRow.worksheet.mergeCells('A2', 'D2');
    worksheet.addRow('');
    const DateToday = this.shared.datePipe.transform(
      new Date(),
      'MM/dd/yyyy h:mm:ss a'
    );
    const DATE_EXTENSION = this.shared.datePipe.transform(new Date(), 'MMddyyyy');
    worksheet.addRow([DateToday]).font = { name: 'Arial', family: 4, size: 9 };
    // const Appointmentdata = this.LoanerBlock.map((_arrayElement: any) =>
    //   Object.assign({}, _arrayElement)
    // );
    const ReportFilter = worksheet.addRow(['Report Filters :']);
    ReportFilter.font = { name: 'Arial', family: 4, size: 10, bold: true };
    const Groups = worksheet.getCell('A6');
    Groups.value = 'Group :';
    Groups.font = { name: 'Arial', family: 4, size: 9, bold: true };
    const groups = worksheet.getCell('B6');
    groups.value =
      this.comm.groupsandstores.filter((val: any) => val.sg_id == this.groupId.toString())[0].sg_name;
    groups.font = { name: 'Arial', family: 4, size: 9 };
    groups.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
    worksheet.mergeCells('B7', 'K9');
    const Stores = worksheet.getCell('A7');
    Stores.value = 'Stores :'
    const stores = worksheet.getCell('B7');
    stores.value = this.ExcelStoreNames == 0
      ? '-'
      : this.ExcelStoreNames == null
        ? '-'
        : this.ExcelStoreNames.toString().replaceAll(',', ', ');
    stores.font = { name: 'Arial', family: 4, size: 9 };
    stores.alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    Stores.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
    };
    const Timeframe = worksheet.addRow(['Timeframe :']);
    Timeframe.getCell(1).font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
    };
    const timeframe = worksheet.getCell('B10');
    timeframe.value = this.FromDate + ' to ' + this.ToDate;
    timeframe.font = { name: 'Arial', family: 4, size: 9 };
    worksheet.addRow('');
    const PresentYear = this.shared.datePipe.transform(this.FromDate, 'yyyy');
    const FromDate = this.shared.datePipe.transform(this.FromDate, 'dd');
    const ToDate = this.shared.datePipe.transform(this.ToDate, 'dd');
    const PresentMonth = this.shared.datePipe.transform(this.FromDate, 'MMM');
    // const mergedHeaderRow = worksheet.addRow([
    //   'Retail Performance', '', '', '', '', '', '', '', '', '', // 10 columns
    //   'Internal Performance', '', '', '', ''                    // 5 columns
    // ]);
    // worksheet.mergeCells(`A11:J11`);
    // worksheet.mergeCells(`K11:O11`);
    // mergedHeaderRow.getCell(1).font = { name: 'Arial', size: 10, bold: true };
    // mergedHeaderRow.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
    // mergedHeaderRow.getCell(11).font = { name: 'Arial', size: 10, bold: true };
    // mergedHeaderRow.getCell(11).alignment = { horizontal: 'center', vertical: 'middle' };
    let dateYear = worksheet.getCell('A11');
    dateYear.value = '';
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
      fgColor: { argb: 'FF1F4E78' },
      bgColor: { argb: 'FF0000FF' },
    };
    dateYear.border = { right: { style: 'dotted' } };
    worksheet.mergeCells('B11', 'J11');
    let units = worksheet.getCell('B11');
    units.value = 'Retail Tire Performance';
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
      fgColor: { argb: 'FF1F4E78' },
      bgColor: { argb: 'FF0000FF' },
    };
    units.border = { right: { style: 'dotted' } };
    worksheet.mergeCells('K11', 'P11');
    let frontgross = worksheet.getCell('K11');
    frontgross.value = 'Internal Tire Performance';
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
      fgColor: { argb: 'FF1F4E78' },
      bgColor: { argb: 'FF0000FF' },
    };
    frontgross.border = { right: { style: 'dotted' } };
    let bindingHeaders: any = []
    bindingHeaders = ['Store_Name', 'Retail_Tire_Sale', 'Retail_Tire_Gross', 'Retail_GP_Per', 'Retail_CarCount', 'Retail_TireGoal', 'Retail_Tires_Sold', 'Retail_Goal_Per', 'Retail_AvgCost', 'Retail_AvgProfit', 'Missed_opportunity', 'Interal_Tire_sale', 'Interal_Tire_Gross', 'Internal_GP_Per', 'Internal_Tires_Sold', 'Internal_AvgCost']
    let headerRow: any = []
    headerRow = [`${PresentMonth}  ${FromDate} - ${ToDate}, ${PresentYear}`, 'Sales $', 'Gross $', 'GP %', 'Car Count', 'Tire Goal', 'Tires Sold', '% of Goal', 'AVG Cost', 'AVG Profit', 'Missed Opportunity', 'Sales $', 'Gross $', 'GP %', 'Tires Sold', 'AVG Cost']
    const currencyFields: any = ['Retail_Tire_Sale', 'Retail_Tire_Gross', 'Retail_AvgCost', 'Retail_AvgProfit', 'Interal_Tire_sale', 'Interal_Tire_Gross', 'Internal_Tires_Sold', 'Internal_AvgCost', 'Missed_opportunity'];
    const percentageFields: any = ['Retail_GP_Per', 'Retail_Goal_Per', 'Internal_GP_Per']
    const excelHeader = worksheet.addRow(headerRow);
    excelHeader.eachCell({ includeEmpty: false }, (cell, colNumber) => {
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF1F4E78' }
      };
      cell.alignment = { horizontal: 'center' };
    });
    for (const info of this.SalesData) {
      const rowData = bindingHeaders.map((key: any) => {
        const val = info[key];
        return val === 0 || val == null ? '-' : val
      });
      const dealerRow = worksheet.addRow(rowData);
      dealerRow.font = { bold: true };
      bindingHeaders.forEach((key: any, index: any) => {
        const cell = dealerRow.getCell(index + 1);
        if (currencyFields.includes(key) && typeof cell.value === 'number') {
          cell.numFmt = '"$"#,##0.00';
          cell.alignment = { horizontal: 'right' };
        } else if (percentageFields.includes(key) && typeof cell.value === 'number') {
          cell.value = cell.value / 100; // Convert 45 to 0.45
          cell.numFmt = '0%';
          cell.alignment = { horizontal: 'right' };
        }
        else if (!isNaN(Number(cell.value))) {
          cell.alignment = { horizontal: 'center' };
        } else {
          cell.alignment = { horizontal: 'center' };
        }
      });
    }
    worksheet.columns.forEach((column: any) => {
      let maxLength = 5;
      column.eachCell({ includeEmpty: true }, (cell: any) => {
        let columnLength = 0;
        if (cell.value != null) {
          const cellText = cell.value.toString();
          columnLength = cellText.length;
        }
        maxLength = Math.max(maxLength, columnLength);
      });
      column.width = 30;
    });
   
     this.shared.exportToExcel(workbook, `Tire Group Sales Reports_${new Date().getTime()}.xlsx`);
  
  }

}