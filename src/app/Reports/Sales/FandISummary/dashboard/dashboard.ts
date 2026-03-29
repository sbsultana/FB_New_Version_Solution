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

  FISummaryData: any = [];
  fiManagersname: any = '';

  NoData: any = '';
  groups: any = [1];
  reportOpenSub!: Subscription;
  reportGetting!: Subscription;
  Pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;
  // Report Popup


  dealtype: any = ['Retail', 'Lease'];
  Keys: any = [
    { name: 'Counts', symbol: 'N', displayname: 'Counts (#)' },
    { name: 'Penetration', symbol: 'P', displayname: 'Penetration (%)' },
    { name: 'Income', symbol: 'D1', displayname: 'Income (Total $)' },
    { name: 'AvgIncome', symbol: 'D0', displayname: 'Income (Average $)' },
  ]

  stores: any = []
  groupsArray: any = [];
  storename: any = ''
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  groupId: any = 0;
  storeIds: any = 0;

  storesFilterData: any = {
    'groupsArray': this.groupsArray, 'groupId': this.groupId, 'storesArray': this.stores, 'storeids': '1', 'type': 'S', 'others': 'N',
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
    const clickedInside = (event.target as HTMLElement).closest('.dropdown-toggle, .reportstores-card , .timeframe');
    if (!clickedInside) {
      this.activePopover = -1;
    }
  }
  constructor(
    public shared: Sharedservice, public setdates: Setdates, private comm: common, private cp: CurrencyPipe, private toast: ToastService,
  ) {
    this.shared.setTitle(this.comm.titleName + '-F & I Summary');



    if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
      this.groupId = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Preferences
      this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids.split(',')[0];
      if (localStorage.getItem('flag') == 'V') {
        this.initializeDates(this.comm.redirectionFrom.Type)
        this.DateType = this.comm.redirectionFrom.Type
      }
      else {
        this.initializeDates('MTD')
        this.DateType = 'MTD'
      }
      this.maxDate = new Date();
      this.minDate = new Date();
      this.minDate.setFullYear(this.maxDate.getFullYear() - 3);
      this.maxDate.setDate(this.maxDate.getDate());
    }
    if (this.shared.common.groupsandstores.length > 0) {
      this.groupsArray = this.shared.common.groupsandstores.filter((val: any) => val.sg_id != this.shared.common.reconID);
      this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
      this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_Name : this.groupName = ''
      this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
      this.getStoresandGroupsValues()
    }
    this.setDates(this.DateType)
    this.setHeaderData()
    this.FandIManagers()
  }
  ngOnInit(): void {
  }
  initializeDates(type: any) {
    let dates: any = this.setdates.setDates(type)
    this.FromDate = dates[0];
    this.ToDate = dates[1];
    console.log(this.FromDate, this.ToDate);
    this.getSalesData();

    localStorage.setItem('time', type);
  }

  getSalesData() {
    if (this.storeIds != undefined && this.storeIds.length > 0 && this.FromDate != '' && this.ToDate != '') {
      console.log(this.FromDate, this.ToDate, '................ From Date and Todate');

      this.shared.spinner.show();
      this.GetData();
    } else {
      // this.shared.spinner.hide();
    }
  }
  GetData() {
    // this.comm.redirectionFrom.flag = 'M'

    const obj = {
      "store": this.storeIds.toString(),
      "startdate": this.FromDate,
      "enddate": this.ToDate,
      "dealtype": this.dealtype.toString(),
      "FIMGR": this.financeManagerId && this.financeManagerId.length > 0 ? this.financeManagerId.toString() : ''
    };
    const curl = environment.apiUrl + 'GetFandIMVPReportV1';
    this.shared.api
      .postmethod(this.comm.routeEndpoint + 'GetFandIMVPReportV1', obj)
      .subscribe(
        (res) => {
          const currentTitle = document.title;
          this.shared.api.logSaving(curl, {}, '', res.message, currentTitle);
          this.FISummaryData = [];
          if (res && res.response && res.response.length > 0) {
            this.FISummaryData = res.response;
            console.log(res.response, '............');
            this.FISummaryData.some(function (x: any) {
              if (x.Penetration != undefined && x.Penetration != '' && x.Penetration != null) {
                x.Penetration = JSON.parse(x.Penetration);
              }
              if (x.Income != undefined && x.Income != '' && x.Income != null) {
                x.Income = JSON.parse(x.Income);
              }
              if (x.AvgIncome != undefined && x.AvgIncome != '' && x.AvgIncome != null) {
                x.AvgIncome = JSON.parse(x.AvgIncome);
              }
              if (x.Counts != undefined && x.Counts != '' && x.Counts != null) {
                x.Counts = JSON.parse(x.Counts);
                x.Counts.some(function (y: any) {
                  if (y.Products != undefined && y.Products != '' && y.Products != null) {
                    y.Products = JSON.parse(y.Products)
                  }
                })
              }
            });
            this.shared.spinner.hide();
            this.NoData = '';
          } else {
            this.shared.spinner.hide();
            this.NoData = 'No Data Found!!';
          }
        },
        (error) => {
          this.toast.show('502 Bad Gate Way Error', 'danger', 'Error');
          this.shared.spinner.hide();
          this.NoData = 'No Data Found';
        }
      );
  }
  financeManager: any = []
  financeManagerId: any = [];
  FandIManagers() {
    const obj = {
      AS_ID: this.storeIds,
      type: 'F',
    }
    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetEmployeesDev', obj).subscribe((res: any) => {
      if (res.status == 200) {
        if (res.response && res.response.length > 0) {
          this.financeManager = res.response.filter((e: any) => e.FiName != 'Unknown');
          this.financeManagerId = this.financeManager.map(function (a: any) { return a.FiId; });
        }
      }
    })
  }

  employees(block: any, e: any) {
    if (block === 'FM') {
      const index = this.financeManagerId.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.financeManagerId.splice(index, 1);
      } else {
        this.financeManagerId.push(e);
      }
      if (this.financeManagerId.length == 1) {
        this.fiManagersname = this.financeManager.filter((val: any) => val.FiId == e)[0].FiName
      }
    }


    if (block === 'AllFM') {
      if (e === 0) {
        this.financeManagerId = this.financeManager.map(
          (fm: any) => fm.FiId
        );
      } else if (e === 1) {
        this.financeManagerId = [];
      }
    }
  }
  async openSalesModal(dealnumber: any, vin: any, storeid: any, stock: any, source: any, custno: any) {
    const module = await import('../../../../Layout/cdpdataview/deal/deal-module');
    const component = module.Deal;

    const modalRef = this.shared.ngbmodal.open(component, { size: 'xl', windowClass: 'connectedmodal' });
    modalRef.componentInstance.data = { dealno: dealnumber, vin: vin, storeid: storeid, stock: stock, source: source, custno: custno }; // Pass data to the modal component    
    modalRef.result.then((result) => {
      // this.topScroll()
      console.log(result); // Handle modal close result
    }, (reason) => {
      // this.topScroll()
      console.log(`Dismissed: ${reason}`); // Handle dismiss reason
    });
  }
  expandDataArray: any = []
  expandorcollapse(cat: any) {
    const index = this.expandDataArray.findIndex((i: any) => i == cat);
    if (index >= 0) {
      this.expandDataArray.splice(index, 1);
    } else {
      this.expandDataArray.push(cat);
    }
  }
  getTotal(newvalue: any, usedvalue: any) {
    if (newvalue && usedvalue)
      return newvalue + usedvalue
    else if (newvalue)
      return newvalue
    else if (usedvalue)
      return usedvalue
    else
      return 0
  }
  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    } else if (value < 0) {
      return false;
    }
    return true;
  }
  popupReference!: NgbModalRef;
  popupvalues: any = { block: '', key: '', saletype: '', Category: '', prodname: '' }
  FISummaryDetails: any = []
  FISummaryDetailsNoData: any = ''
  detailpopup(tmp: any, data: any, block: any, key: any, prodName?: any) {
    this.popupvalues.block = block;
    this.popupvalues.key = key;
    this.popupvalues.saletype = data.Category
    this.popupvalues.Category = data.Category
    this.popupvalues.prodname = prodName ? prodName : ''
    this.FISummaryDetailsNoData = '';
    this.FISummaryDetails = []
    this.popupReference = this.shared.ngbmodal.open(tmp, { size: 'xl', backdrop: 'static', keyboard: true, centered: true, modalDialogClass: 'custom-modal' })
    this.GetFISumDetails(data, block, prodName)
  }
  closePopup() {
    if (this.popupReference) {
      this.popupReference.close();
    }
  }
  GetFISumDetails(data: any, block: any, ProductName: any) {
    console.log(data.Category);

    this.popupvalues.saletype = data.Category != 'Cash' && data.Category != 'Finance' && data.Category != 'Lease' ? '' : data.Category
    const obj = {
      "PT_Type": data.PT_ID,
      "store": this.storeIds.toString(),
      "startdate": this.FromDate,
      "enddate": this.ToDate,
      "dealtype": block == 'Total' ? '' : block,
      "saletype": data.Category == 'Total Units' ? 'Cash,Finance,Lease' : (data.Category != 'Cash' && data.Category != 'Finance' && data.Category != 'Lease' ? '' : data.Category),
      "ProdName": ProductName ? ProductName : ''

    }
    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetFandIMVPReportV1Details', obj).subscribe(
      (res) => {
        if (res.response && res.response.length > 0) {
          this.FISummaryDetails = res.response
          this.FISummaryDetailsNoData = ''
        } else {
          this.FISummaryDetailsNoData = 'No Data Found'
        }
      },
      (error) => {
        this.FISummaryDetailsNoData = 'No Data Found'
      })
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
  ngAfterViewInit() {

    this.shared.api.getStores().subscribe((res: any) => {
      if (this.comm.pageName == 'F & I Summary') {
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
        if (res.obj.title == 'F & I Summary') {
          if (res.obj.state == true) {
            this.exportToExcel();
          }
        }
      }
    });
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res) => {
      if (this.email != undefined) {
        if (res.obj.title == 'F & I Summary') {
          if (res.obj.stateEmailPdf == true) {
            // this.sendEmailData(res.obj.Email, res.obj.notes, res.obj.from);
          }
        }
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res) => {
      if (this.print != undefined) {
        if (res.obj.title == 'F & I Summary') {
          if (res.obj.statePrint == true) {
            // this.GetPrintData();
          }
        }
      }
    });
    this.Pdf = this.shared.api.getExportToPDFAllReports().subscribe((res) => {
      if (this.Pdf != undefined) {
        console.log('PDF');
        if (res.obj.title == 'F & I Summary') {
          console.log('PDF F & I Summary');
          if (res.obj.statePDF == true) {
            // this.generatePDF();
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
    this.FandIManagers()
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
      'type': 'S', 'others': 'N'
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


  viewDeal(dealData: any) {
    // const modalRef = this.shared.ngbmodal.open(DealRecapComponent, { size: 'md', windowClass: 'connectedmodal' });
    // modalRef.componentInstance.data = { dealno: dealData.ad_dealid, storeid: dealData.Store_Id, stock: dealData.Stock, vin: dealData.VIN, custno: dealData?.customernumber }; // Pass data to the modal component    
    // modalRef.result.then((result) => {
    //   console.log(result); // Handle modal close result
    // }, (reason) => {
    //   console.log(`Dismissed: ${reason}`); // Handle dismiss reason
    // });
  }
  setHeaderData() {
    const headerdata = {
      title: 'F & I Summary',
      stores: this.storeIds,
      fromdate: this.FromDate,
      todate: this.ToDate,
      groups: this.groups,
      dealtype: this.dealtype,
      fimanagers: this.financeManagerId,
      datetype: this.DateType
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
  activePopover: number = -1;
  togglePopover(popoverIndex: number) {
    if (this.activePopover === popoverIndex) {
      this.activePopover = -1;
    } else {
      this.activePopover = popoverIndex;
    }
  }

  FIManagerName: any = ''
  multipleorsingle(block: any, e: any) {
    if (block == 'FI') {
      if (e == 'All') {
        if (this.financeManagerId.length == this.financeManager.length) {
          this.financeManagerId = []
        } else {
          this.financeManagerId = this.financeManager.map(function (a: any) {
            return a.FiId;
          });
        }
      }
      else {
        const index = this.financeManagerId.findIndex((i: any) => i == e);
        if (index >= 0) {
          this.financeManagerId.splice(index, 1);
        } else {
          this.financeManagerId.push(e);
          if (this.financeManagerId.length == 1) {
            this.FIManagerName = this.financeManager.filter((val: any) => val.FiId == e)[0].FiName
          }
        }
      }
    }
    if (block == 'DL') {
      const index = this.dealtype.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.dealtype.splice(index, 1);
      } else {
        this.dealtype.push(e);
      }
    }
  }
  viewreport() {
    this.activePopover = -1
    if (this.storeIds.length == 0) {
      this.toast.show('Please Select Atleast One Store', 'warning', 'Warning');
    }
    else if (this.dealtype == undefined || this.dealtype.length == 0) {
      this.toast.show('Please Select Atleast One Deal Type', 'warning', 'Warning');
    }
    else {
      this.setHeaderData();
      this.getSalesData()
    }
  }

  ExcelStoreNames: any = []
  managers: any = []
  exportToExcel(): void {
    let storeNames: any = [];
    let store = this.storeIds
    storeNames = this.comm.groupsandstores
      .find((v: any) => v.sg_id == this.groups)
      ?.Stores.filter((item: any) => store.includes(Number(item.ID)));
    if (store.length == this.comm.groupsandstores.filter((v: any) => v.sg_id == this.groups)[0].Stores.length) {
      this.ExcelStoreNames = 'All Stores'
    } else {
      this.ExcelStoreNames = storeNames.map(function (a: any) {
        return a.storename;
      });
      console.log(this.ExcelStoreNames);

    }
    this.managers = this.financeManager.filter((item: any) => this.financeManagerId.includes(item.FiId)).map((item: any) => item.FiName)
    console.log(this.ExcelStoreNames.toString().replaceAll(',', ', '), this.financeManager.filter((item: any) => this.financeManagerId.includes(item.FiId)).map((item: any) => item.FiName));

    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('F & I Summary');
    worksheet.views = [
      {
        // state: 'frozen',
        // ySplit: 1,
        // topLeftCell: 'A2',
        // // showGridLines: false,

        state: 'frozen',
        ySplit: 17, // Number of rows to freeze (2 means the first two rows are frozen)
        topLeftCell: 'A18', // Specify the cell to start freezing from (in this case, the third row)
        showGridLines: true,
      },
    ];
    worksheet.addRow('');
    const titleRow = worksheet.addRow(['F and I Summary']);
    titleRow.eachCell((cell, number) => {
      cell.alignment = { indent: 1, vertical: 'top', horizontal: 'center' };
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

    const ReportFilter = worksheet.addRow(['Report Filters :']);
    ReportFilter.font = { name: 'Arial', family: 4, size: 10, bold: true };

    const Groups = worksheet.getCell('A6');
    Groups.value = 'Group :';
    Groups.font = { name: 'Arial', family: 4, size: 9, bold: true };
    const groups = worksheet.getCell('B6');
    groups.value =
      this.comm.groupsandstores.filter((val: any) => val.sg_id == this.groups.toString())[0].sg_name;
    groups.font = { name: 'Arial', family: 4, size: 9 };
    groups.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };

    worksheet.mergeCells('B7', 'K9');
    const Stores = worksheet.getCell('A7');
    Stores.value = 'Store :'
    const stores = worksheet.getCell('B7');
    stores.value = this.ExcelStoreNames == 0 ? '-' : this.ExcelStoreNames == null ? '-' : this.ExcelStoreNames.toString().replaceAll(',', ', ');
    stores.font = { name: 'Arial', family: 4, size: 9 };
    stores.alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    Stores.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
    };



    worksheet.mergeCells('B11:K13');

    const financeManagerId = worksheet.getCell('A11');
    financeManagerId.value = 'F & I Managers :';
    financeManagerId.font = { name: 'Arial', family: 4, size: 9, bold: true };
    financeManagerId.alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    const fimangers = worksheet.getCell('B11');
    fimangers.value = this.financeManager
      .filter((item: any) => this.financeManagerId.includes(item.FiId))
      .map((item: any) => item.FiName)
      .toString();
    fimangers.font = { name: 'Arial', family: 4, size: 9, bold: false };
    fimangers.alignment = { vertical: 'top', horizontal: 'left', wrapText: false };



    const Timeframe = worksheet.addRow(['Time Frame :']);
    Timeframe.getCell(1).font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
    };
    const timeframe = worksheet.getCell('B14');
    timeframe.value = this.FromDate + ' to ' + this.ToDate;
    timeframe.font = { name: 'Arial', family: 4, size: 9 };

    worksheet.addRow('');
    const PresentYear = this.shared.datePipe.transform(this.FromDate, 'yyyy');
    const FromDate = this.shared.datePipe.transform(this.FromDate, 'dd');
    const ToDate = this.shared.datePipe.transform(this.ToDate, 'dd');
    const PresentMonth = this.shared.datePipe.transform(this.FromDate, 'MMM');

    worksheet.addRow('');
    const bindingHeaders = [
      'Category', 'New', 'Used', 'Total',
    ];
    const subheader = ['ProductName', 'New', 'Used', 'Total',]
    const currencyFields: any = [];
    const headerRow = [
      `${PresentMonth}  ${FromDate} - ${ToDate}, ${PresentYear}`, 'New', 'Used', 'Total'
    ];
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
    for (const key of this.Keys) {
      const rowData = [key.displayname, '', '', '']
      const dealerRow = worksheet.addRow(rowData);
      dealerRow.font = { bold: true };
      for (const info of this.FISummaryData) {
        for (const dept of info[key.name]) {
          const nestedRowData = bindingHeaders.map(keys => {
            const val = dept[keys];
            return val === 0 || val == null ? '-' : val;
          });
          const nestedRow = worksheet.addRow(nestedRowData);
          bindingHeaders.forEach((keys, index) => {
            const cell = nestedRow.getCell(index + 1);
            if (typeof cell.value === 'number') {
              if (key.symbol == 'P') {
                cell.numFmt = '#,##0.0';
              } else {
                cell.numFmt = '#,##0';
              }
              cell.alignment = { horizontal: 'right' };
            } else if (!isNaN(Number(cell.value))) {
              cell.alignment = { horizontal: 'right' };
            }
          });
          if (dept.Products != '' && dept.Products != null && dept.Products != undefined) {
            for (const pd of dept.Products) {
              const nestedRowData = subheader.map(keys => {
                const val = pd[keys];
                return val === 0 || val == null ? '-' : keys == 'ProductName' ? ('         ' + val) : val;
              });
              const nestedRow = worksheet.addRow(nestedRowData);
              subheader.forEach((keys, index) => {
                const cell = nestedRow.getCell(index + 1);
                if (typeof cell.value === 'number') {
                  if (key.symbol == 'P') {
                    cell.numFmt = '#,##0.0';
                  } else {
                    cell.numFmt = '#,##0';
                  }
                  cell.alignment = { horizontal: 'right' };
                } else if (!isNaN(Number(cell.value))) {
                  cell.alignment = { horizontal: 'right' };
                }
              });
            }
          }
        }
      }
    }
    worksheet.getColumn(1).width = 40;
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
    workbook.xlsx.writeBuffer().then((buffer: any) => {
      const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      // FileSaver.saveAs(blob, 'Loaner Inventory_' + DATE_EXTENSION + EXCEL_EXTENSION);
      this.shared.exportToExcel(workbook, 'F & I Summary');
      // FileSaver.saveAs(blob, `F & I Summary_${new Date().getTime()}.xlsx`);
    });
  }
  excelDownload(): void {
    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('F & I Summary Details');
    worksheet.views = [
      {
        state: 'frozen',
        ySplit: 1,
        topLeftCell: 'A2',
        // showGridLines: false,
      },
    ];
    let bindingHeaders: any = []
    this.popupvalues.saletype == '' ? bindingHeaders = ['Date', 'ad_dealid', 'Stock', 'VIN', 'CustomerName', 'FIManager', 'ad_frontgross', 'ad_backgross', 'TotalGross'] :
      bindingHeaders = ['Date', 'ad_dealid', 'Stock', 'CustomerName', 'FIManager', 'ad_frontgross', 'ad_backgross', 'TotalGross']
    let headerRow: any = []
    this.popupvalues.saletype == '' ? headerRow = ['Date', 'Deal #', 'Stock', 'VIN', 'Customer Name', 'FI Manager', 'Sale', 'Cost', 'Total Gross'] :
      headerRow = ['Date', 'Deal #', 'Stock', 'Customer Name', 'FI Manager', 'Front Gross', 'Back Gross', 'Total Gross']
    const currencyFields: any = ['ad_frontgross', 'ad_backgross', 'TotalGross'];
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
    for (const info of this.FISummaryDetails) {
      const rowData = bindingHeaders.map((key: any) => {
        const val = info[key];
        return val === 0 || val == null ? '-' : key == 'Date' ? this.shared.datePipe.transform(val, 'MM.dd.YYYY') : val
      });
      const dealerRow = worksheet.addRow(rowData);
      dealerRow.font = { bold: true };
      bindingHeaders.forEach((key: any, index: any) => {
        const cell = dealerRow.getCell(index + 1);
        if (currencyFields.includes(key) && typeof cell.value === 'number') {
          cell.numFmt = '"$"#,##0.00';
          cell.alignment = { horizontal: 'right' };
        } else if (!isNaN(Number(cell.value))) {
          cell.alignment = { horizontal: 'right' };
        }
      });
    }
    worksheet.columns.forEach((column: any) => {
      let maxLength = 15;
      column.eachCell({ includeEmpty: true }, (cell: any) => {
        let columnLength = 0;
        if (cell.value != null) {
          const cellText = cell.value.toString();
          columnLength = cellText.length;
        }
        maxLength = Math.max(maxLength, columnLength);
      });
      column.width = maxLength + 2;
    });
    workbook.xlsx.writeBuffer().then((buffer: any) => {
      const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      this.shared.exportToExcel(workbook, 'F & I Summary Details');

    });
  }
  // GetPrintData() {
  //   window.print();
  // }
  // generatePDF() {
  //   this.shared.spinner.show();
  //   const printContents = document.getElementById('SellingGross')!.innerHTML;
  //   const iframe = document.createElement('iframe');
  //   iframe.style.position = 'absolute';
  //   iframe.style.width = '0px';
  //   iframe.style.height = '0px';
  //   iframe.style.border = 'none';
  //   document.body.appendChild(iframe);
  //   const doc = iframe.contentWindow?.document;
  //   if (!doc) {
  //     console.error('Failed to create iframe document');
  //     return;
  //   }
  //   doc.open();
  //   doc.write(`
  //   <html>
  //           <head>
  //             <title>F & I Summary</title>
  //                <style>
  //                @font-face {
  //                 font-family: 'GothamBookRegular';
  //                 src: url('assets/fonts/Gotham\ Book\ Regular.otf') format('otf'), /* Chrome 6+, Firefox 3.6+, IE 9+, Safari 5.1+ */
  //                      url('assets/fonts/Gotham\ Book\ Regular.otf') format('opentype'); /* Chrome 4+, Firefox 3.5, Opera 10+, Safari 3—5 */
  //               }
  //               @font-face {
  //                 font-family: 'Roboto';
  //                 src: url('assets/fonts/Roboto-Regular.ttf') format('ttf'), /* Chrome 6+, Firefox 3.6+, IE 9+, Safari 5.1+ */
  //                      url('assets/fonts/Roboto-Regular.ttf') format('truetype'); /* Chrome 4+, Firefox 3.5, Opera 10+, Safari 3—5 */
  //               }
  //               @font-face {
  //                 font-family: 'RobotoBold';
  //                 src: url('assets/fonts/Roboto-Bold.ttf') format('ttf'), /* Chrome 6+, Firefox 3.6+, IE 9+, Safari 5.1+ */
  //                      url('assets/fonts/Roboto-Bold.ttf') format('truetype'); /* Chrome 4+, Firefox 3.5, Opera 10+, Safari 3—5 */
  //               }
  //               .negative {
  //                 color: red;
  //               }
  //               .HideItems{
  //                 display:none;
  //                }            
  // .backgrossbtn {
  //                 float: left;
  //                 font-size: 11px;
  //                 background-color: #2d91f1;
  //                 color: white;
  //                 padding: 4px;
  //                 border: 1px solid #2d91f1;
  //                 cursor: pointer;
  //                 border-radius: 3px;
  //             }
  //               .Selectedrow {
  //                 background-color: #d3ecff !important;
  //                 color: #000 !important;
  //             }
  //               .justify-content-between {
  //                 justify-content: space-between !important;
  //             }
  //             .d-flex {
  //                 display: flex !important;
  //             }
  //               .bg-white {
  //                 background: #ffffff !important;
  //             }
  //               .negative {
  //                 color: red;
  //               }
  //               .negative {
  //                 color: red;
  //               }
  //               .performance-scorecard .table>:not(:first-child) {
  //                 border-top: 0px solid #ffa51a
  //               }
  //               .performance-scorecard .table {
  //                 text-align: center;
  //                 text-transform: capitalize;
  //                 border: transparent;
  //                 width: 100%;
  //               }
  //               .performance-scorecard .table th,
  //               .performance-scorecard .table td{
  //                 white-space: nowrap;
  //                 vertical-align: top;
  //               }
  //               .performance-scorecard .table th:first-child,
  //               .performance-scorecard .table td:first-child {
  //                 position: sticky;
  //                 left: 0;
  //                 z-index: 1;
  //                 background-color: #337ab7;
  //               }
  //               .performance-scorecard .table tr:nth-child(odd) td:first-child,
  //               .performance-scorecard .table tr:nth-child(odd) td:nth-child(2) {
  //                 // background-color: #e9ecef;
  //                 background-color: #ffffff;
  //               }
  //               .performance-scorecard .table tr:nth-child(even) td:first-child,
  //               .performance-scorecard .table tr:nth-child(even) td:nth-child(2) {
  //                 background-color: #ffffff;
  //               }
  //               .performance-scorecard .table tr:nth-child(odd) {
  //                 // background-color: #e9ecef;
  //                 background-color: #ffffff;
  //               }
  //               .performance-scorecard .table tr:nth-child(even) {
  //                 background-color: #ffffff;
  //               }
  //               .performance-scorecard .table .spacer {
  //                 // width: 50px !important;
  //                 background-color: #cfd6de !important;
  //                 border-left: 1px solid #cfd6de !important;
  //                 border-bottom: 1px solid #cfd6de !important;
  //                 border-top: 1px solid #cfd6de !important;
  //                }
  //               .performance-scorecard .table .hidden {
  //                 display: none !important;
  //               }
  //               .performance-scorecard .table .bdr-rt {
  //                 border-right: 1px solid #abd0ec;
  //               }
  //               .performance-scorecard .table thead {
  //                 position: sticky;
  //                 top: 0;
  //                 z-index: 99;
  //                 font-family: 'FaktPro-Bold';
  //                 font-size: 0.8rem;
  //               }
  //               .performance-scorecard .table thead th {
  //                 padding: 5px 10px;
  //                 margin: 0px;
  //               }
  //               .performance-scorecard .table thead .bdr-btm {
  //                 border-bottom: #005fa3;
  //               }
  //               .performance-scorecard .table thead tr:nth-child(1) {
  //                 background-color: #fff !important;
  //                 color: #000;
  //                 text-transform: uppercase;
  //                 border-bottom: #cfd6de;
  //               }
  //               .performance-scorecard .table thead tr:nth-child(2) {
  //                 background-color: #fff !important;
  //                 color: #000;
  //                 text-transform: uppercase;
  //                 border-bottom: #cfd6de;
  //                 box-shadow: inset 0 1px 0 0 #cfd6de;
  //               }
  //               .performance-scorecard .table thead tr:nth-child(2) th :nth-child(1) {
  //                 background-color: #337ab7 !important;
  //                 color: #fff;
  //               }
  //               .performance-scorecard .table thead tr:nth-child(3) {
  //                 background-color: #fff !important;
  //                 color: #000;
  //                 text-transform: uppercase;
  //                 border-bottom: #cfd6de;
  //                 box-shadow: inset 0 1px 0 0 #cfd6de;
  //               }
  //               .performance-scorecard .table thead tr:nth-child(3) th :nth-child(1) {
  //                 background-color: #337ab7 !important;
  //                 color: #fff;
  //               }
  //               .performance-scorecard .table tbody {
  //                 font-family: 'FaktPro-Normal';
  //                 font-size: .9rem;
  //               }
  //               .performance-scorecard .table tbody td {
  //                 padding: 2px 10px;
  //                 margin: 0px;
  //                 border: 1px solid #cfd6de
  //               }
  //               .performance-scorecard .table tbody tr {
  //                 border-bottom: 1px solid #37a6f8;
  //                 border-left: 1px solid #37a6f8
  //               }
  //               .performance-scorecard .table tbody td:first-child {
  //                 text-align: start;
  //                 box-shadow: inset -1px 0 0 0 #cfd6de;
  //               }
  //               .performance-scorecard .table tbody tr td:not(:first-child) {
  //                 text-align: right !important;
  //               }
  //               .performance-scorecard .table tbody .sub-title {
  //                 font-size: .8rem !important;
  //               }
  //               .performance-scorecard .table tbody .sub-subtitle{
  //                 font-size: .7rem !important;
  //               }
  //               .performance-scorecard .table tbody .text-bold {
  //                 font-family: 'FaktPro-Bold';
  //               }
  //               .performance-scorecard .table tbody .darkred-bg {
  //                 background-color: #282828 !important;
  //                 color: #fff;
  //               }
  //               .performance-scorecard .table tbody .lightblue-bg {
  //                 background-color: #646e7a !important;
  //                 color: #fff;
  //               }
  //               .performance-scorecard .table tbody .gold-bg {
  //                 background-color: #ffa51a;
  //                 color: #fff;
  //               }
  //                </style>
  //           </head>
  //       <body id='content'>
  //       ${printContents}
  //       </body>
  //         </html>`);
  //   doc.close();
  //   const div = doc.getElementById('content');
  //   const options = {
  //     logging: true,
  //     allowTaint: false,
  //     useCORS: true,
  //   };
  //   if (!div) {
  //     console.error('Element not found');
  //     return;
  //   }
  //   html2canvas(div, options)
  //     .then((canvas) => {
  //       let imgWidth = 285;
  //       let pageHeight = 206;
  //       let imgHeight = (canvas.height * imgWidth) / canvas.width;
  //       let heightLeft = imgHeight;
  //       const contentDataURL = canvas.toDataURL('image/png');
  //       let pdfData = new jsPDF('l', 'mm', 'a4', true);
  //       let position = 5;
  //       function addExtraDataToPage(pdf: any, extraData: any, positionY: any) {
  //         pdf.text(extraData, 10, positionY);
  //       }
  //       function addPageAndImage(pdf: any, contentDataURL: any, position: any) {
  //         pdf.addPage();
  //         pdf.addImage(
  //           contentDataURL,
  //           'PNG',
  //           5,
  //           position,
  //           imgWidth,
  //           imgHeight,
  //           undefined,
  //           'FAST'
  //         );
  //       }
  //       pdfData.addImage(
  //         contentDataURL,
  //         'PNG',
  //         5,
  //         position,
  //         imgWidth,
  //         imgHeight,
  //         undefined,
  //         'FAST'
  //       );
  //       addExtraDataToPage(pdfData, '', position + imgHeight + 10);
  //       heightLeft -= pageHeight;
  //       while (heightLeft >= 0) {
  //         position = heightLeft - imgHeight;
  //         addPageAndImage(pdfData, contentDataURL, position);
  //         addExtraDataToPage(pdfData, '', position + imgHeight + 10);
  //         heightLeft -= pageHeight;
  //       }
  //       return pdfData;
  //     })
  //     .then((doc) => {
  //       doc.save('sellingreport.pdf');
  //       // popupWin.close();
  //       this.shared.spinner.hide();
  //     });
  // }
  // sendEmailData(Email: any, notes: any, from: any) {
  //   this.shared.spinner.show();
  //   const printContents = document.getElementById('SellingGross')!.innerHTML;
  //   const iframe = document.createElement('iframe');
  //   // Make the iframe invisible
  //   iframe.style.position = 'absolute';
  //   iframe.style.width = '0px';
  //   iframe.style.height = '0px';
  //   iframe.style.border = 'none';
  //   document.body.appendChild(iframe);
  //   const doc = iframe.contentWindow?.document;
  //   if (!doc) {
  //     console.error('Failed to create iframe document');
  //     return;
  //   }
  //   doc.open();
  //   doc.write(`
  //         <html>
  //             <head>
  //                 <title>F & I Summary</title>
  //                 <style>
  //                 @font-face {
  //                   font-family: 'GothamBookRegular';
  //                   src: url('assets/fonts/Gotham\ Book\ Regular.otf') format('otf'), /* Chrome 6+, Firefox 3.6+, IE 9+, Safari 5.1+ */
  //                        url('assets/fonts/Gotham\ Book\ Regular.otf') format('opentype'); /* Chrome 4+, Firefox 3.5, Opera 10+, Safari 3—5 */
  //                 }
  //                 @font-face {
  //                   font-family: 'Roboto';
  //                   src: url('assets/fonts/Roboto-Regular.ttf') format('ttf'), /* Chrome 6+, Firefox 3.6+, IE 9+, Safari 5.1+ */
  //                        url('assets/fonts/Roboto-Regular.ttf') format('truetype'); /* Chrome 4+, Firefox 3.5, Opera 10+, Safari 3—5 */
  //                 }
  //                 @font-face {
  //                   font-family: 'RobotoBold';
  //                   src: url('assets/fonts/Roboto-Bold.ttf') format('ttf'), /* Chrome 6+, Firefox 3.6+, IE 9+, Safari 5.1+ */
  //                        url('assets/fonts/Roboto-Bold.ttf') format('truetype'); /* Chrome 4+, Firefox 3.5, Opera 10+, Safari 3—5 */
  //                 }
  //                 .negative {
  //                   color: red;
  //                 }
  //                 .HideItems{
  //                   display:none;
  //                  }            
  // .backgrossbtn {
  //                   float: left;
  //                   font-size: 11px;
  //                   background-color: #2d91f1;
  //                   color: white;
  //                   padding: 4px;
  //                   border: 1px solid #2d91f1;
  //                   cursor: pointer;
  //                   border-radius: 3px;
  //               }
  //                 .Selectedrow {
  //                   background-color: #d3ecff !important;
  //                   color: #000 !important;
  //               }
  //                 .justify-content-between {
  //                   justify-content: space-between !important;
  //               }
  //               .d-flex {
  //                   display: flex !important;
  //               }
  //                 .bg-white {
  //                   background: #ffffff !important;
  //               }
  //                 .negative {
  //                   color: red;
  //                 }
  //                 .negative {
  //                   color: red;
  //                 }
  //                 .performance-scorecard .table>:not(:first-child) {
  //                   border-top: 0px solid #ffa51a
  //                 }
  //                 .performance-scorecard .table {
  //                   text-align: center;
  //                   text-transform: capitalize;
  //                   border: transparent;
  //                   width: 100%;
  //                 }
  //                 .performance-scorecard .table th,
  //                 .performance-scorecard .table td{
  //                   white-space: nowrap;
  //                   vertical-align: top;
  //                 }
  //                 .performance-scorecard .table th:first-child,
  //                 .performance-scorecard .table td:first-child {
  //                   position: sticky;
  //                   left: 0;
  //                   z-index: 1;
  //                   background-color: #337ab7;
  //                 }
  //                 .performance-scorecard .table tr:nth-child(odd) td:first-child,
  //                 .performance-scorecard .table tr:nth-child(odd) td:nth-child(2) {
  //                   // background-color: #e9ecef;
  //                   background-color: #ffffff;
  //                 }
  //                 .performance-scorecard .table tr:nth-child(even) td:first-child,
  //                 .performance-scorecard .table tr:nth-child(even) td:nth-child(2) {
  //                   background-color: #ffffff;
  //                 }
  //                 .performance-scorecard .table tr:nth-child(odd) {
  //                   // background-color: #e9ecef;
  //                   background-color: #ffffff;
  //                 }
  //                 .performance-scorecard .table tr:nth-child(even) {
  //                   background-color: #ffffff;
  //                 }
  //                 .performance-scorecard .table .spacer {
  //                   // width: 50px !important;
  //                   background-color: #cfd6de !important;
  //                   border-left: 1px solid #cfd6de !important;
  //                   border-bottom: 1px solid #cfd6de !important;
  //                   border-top: 1px solid #cfd6de !important;
  //                  }
  //                 .performance-scorecard .table .hidden {
  //                   display: none !important;
  //                 }
  //                 .performance-scorecard .table .bdr-rt {
  //                   border-right: 1px solid #abd0ec;
  //                 }
  //                 .performance-scorecard .table thead {
  //                   position: sticky;
  //                   top: 0;
  //                   z-index: 99;
  //                   font-family: 'FaktPro-Bold';
  //                   font-size: 0.8rem;
  //                 }
  //                 .performance-scorecard .table thead th {
  //                   padding: 5px 10px;
  //                   margin: 0px;
  //                 }
  //                 .performance-scorecard .table thead .bdr-btm {
  //                   border-bottom: #005fa3;
  //                 }
  //                 .performance-scorecard .table thead tr:nth-child(1) {
  //                   background-color: #fff !important;
  //                   color: #000;
  //                   text-transform: uppercase;
  //                   border-bottom: #cfd6de;
  //                 }
  //                 .performance-scorecard .table thead tr:nth-child(2) {
  //                   background-color: #fff !important;
  //                   color: #000;
  //                   text-transform: uppercase;
  //                   border-bottom: #cfd6de;
  //                   box-shadow: inset 0 1px 0 0 #cfd6de;
  //                 }
  //                 .performance-scorecard .table thead tr:nth-child(2) th :nth-child(1) {
  //                   background-color: #337ab7 !important;
  //                   color: #fff;
  //                 }
  //                 .performance-scorecard .table thead tr:nth-child(3) {
  //                   background-color: #fff !important;
  //                   color: #000;
  //                   text-transform: uppercase;
  //                   border-bottom: #cfd6de;
  //                   box-shadow: inset 0 1px 0 0 #cfd6de;
  //                 }
  //                 .performance-scorecard .table thead tr:nth-child(3) th :nth-child(1) {
  //                   background-color: #337ab7 !important;
  //                   color: #fff;
  //                 }
  //                 .performance-scorecard .table tbody {
  //                   font-family: 'FaktPro-Normal';
  //                   font-size: .9rem;
  //                 }
  //                 .performance-scorecard .table tbody td {
  //                   padding: 2px 10px;
  //                   margin: 0px;
  //                   border: 1px solid #cfd6de
  //                 }
  //                 .performance-scorecard .table tbody tr {
  //                   border-bottom: 1px solid #37a6f8;
  //                   border-left: 1px solid #37a6f8
  //                 }
  //                 .performance-scorecard .table tbody td:first-child {
  //                   text-align: start;
  //                   box-shadow: inset -1px 0 0 0 #cfd6de;
  //                 }
  //                 .performance-scorecard .table tbody tr td:not(:first-child) {
  //                   text-align: right !important;
  //                 }
  //                 .performance-scorecard .table tbody .sub-title {
  //                   font-size: .8rem !important;
  //                 }
  //                 .performance-scorecard .table tbody .sub-subtitle{
  //                   font-size: .7rem !important;
  //                 }
  //                 .performance-scorecard .table tbody .text-bold {
  //                   font-family: 'FaktPro-Bold';
  //                 }
  //                 .performance-scorecard .table tbody .darkred-bg {
  //                   background-color: #282828 !important;
  //                   color: #fff;
  //                 }
  //                 .performance-scorecard .table tbody .lightblue-bg {
  //                   background-color: #646e7a !important;
  //                   color: #fff;
  //                 }
  //                 .performance-scorecard .table tbody .gold-bg {
  //                   background-color: #ffa51a;
  //                   color: #fff;
  //                 }
  //                 </style>
  //             </head>
  //             <body id='content'>
  //                 ${printContents}
  //             </body>
  //         </html>
  //     `);
  //   doc.close();
  //   const div = doc.getElementById('content');
  //   if (!div) {
  //     console.error('Element not found');
  //     return;
  //   }
  //   const options = {
  //     logging: true,
  //     allowTaint: false,
  //     useCORS: true,
  //   };
  //   html2canvas(div, options)
  //     .then((canvas) => {
  //       let imgWidth = 285;
  //       let pageHeight = 206;
  //       let imgHeight = (canvas.height * imgWidth) / canvas.width;
  //       let heightLeft = imgHeight;
  //       const contentDataURL = canvas.toDataURL('image/png');
  //       let pdfData = new jsPDF('l', 'mm', 'a4', true);
  //       let position = 5;
  //       pdfData.addImage(
  //         contentDataURL,
  //         'PNG',
  //         5,
  //         position,
  //         imgWidth,
  //         imgHeight,
  //         undefined,
  //         'FAST'
  //       );
  //       heightLeft -= pageHeight;
  //       while (heightLeft >= 0) {
  //         position = heightLeft - imgHeight;
  //         pdfData.addPage();
  //         pdfData.addImage(
  //           contentDataURL,
  //           'PNG',
  //           5,
  //           position,
  //           imgWidth,
  //           imgHeight,
  //           undefined,
  //           'FAST'
  //         );
  //         heightLeft -= pageHeight;
  //       }
  //       const pdfBlob = pdfData.output('blob');
  //       const pdfFile = this.blobToFile(pdfBlob, 'SalesGross.pdf');
  //       const formData = new FormData();
  //       formData.append('to_email', Email);
  //       formData.append('subject', 'F & I Summary');
  //       formData.append('file', pdfFile);
  //       formData.append('notes', notes);
  //       formData.append('from', from);
  //       this.shared.api.postmethod(this.comm.routeEndpoint + 'mail', formData).subscribe(
  //         (res: any) => {
  //           console.log('Response:', res);
  //           if (res.status === 200) {
  //             // alert(res.response);
  //             this.toast.success(res.response);
  //           } else {
  //             alert('Invalid Details');
  //           }
  //         },
  //         (error) => {
  //           console.error('Error:', error);
  //         }
  //       );
  //     })
  //     .catch((error) => {
  //       console.error('html2canvas error:', error);
  //     })
  //     .finally(() => {
  //       this.shared.spinner.hide();
  //       // popupWin.close();
  //     });
  // }
  // public blobToFile = (theBlob: Blob, fileName: string): File => {
  //   return new File([theBlob], fileName, {
  //     lastModified: new Date().getTime(),
  //     type: theBlob.type,
  //   });
  // };
}