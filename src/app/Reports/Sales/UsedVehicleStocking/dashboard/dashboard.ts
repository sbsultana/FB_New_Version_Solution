import { Component, OnInit, ViewChild, ElementRef, HostListener } from '@angular/core';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { Setdates } from '../../../../Core/Providers/SetDates/setdates';
import { BsDatepickerConfig, BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { common } from '../../../../common';
import { Stores } from '../../../../CommonFilters/stores/stores';
import { DateRangePicker } from '../../../../CommonFilters/date-range-picker/date-range-picker';
import { CurrencyPipe } from '@angular/common';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
import { Subscription } from 'rxjs';
import { NgbModalRef } from '@ng-bootstrap/ng-bootstrap';
// import { SalesgrossDetailsComponent } from '../../Sales/Gross/salesgross-details/salesgross-details.component';
const EXCEL_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
const EXCEL_EXTENSION = '.xlsx';

@Component({
  selector: 'app-dashboard',
  imports: [SharedModule, BsDatepickerModule, Stores],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
})
export class Dashboard {

  InventoryData: any = [];
  NoInventoryData: any = '';

  groups: any = [1];

  Pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;
  // Report Popup

  DupDateType: any = 'MTD'
  lastyear!: Date;
  ShowHideCI: any = 'Show';
  CIBlock: any = 'Show'
  blocks: any = [
    { label: 'All', value: 'All' },
    { label: 'Top 10', value: '10' },
    { label: 'Top 20', value: '20' },
    { label: 'Top 50', value: '50' },
  ]
  selectedBlock: any = []
  Updatecount: any = 0;
  FromDate: any;
  ToDate: any;
  DateType: any = 'MTD'
  custom: any;

  stores: any = []
  groupsArray: any = [];
  storename: any = ''
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  groupId: any = 1;
  storeIds: any = 0;

  storesFilterData: any = {
    'groupsArray': this.groupsArray, 'groupId': this.groupId, 'storesArray': this.stores, 'storeids': '1', 'type': 'M', 'others': 'N',
    'groupName': this.groupName, 'storename': this.storename, storecount: null, 'storedisplayname': this.storedisplayname
  };


  month!: Date;
  DuplicatDate!: Date;
  minDate!: Date;
  maxDate!: Date;
  bsConfig: Partial<BsDatepickerConfig> = {
    dateInputFormat: 'MMMM/YYYY',
    minMode: 'month'
  };

  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent) {
    const clickedInside = (event.target as HTMLElement).closest('.dropdown-toggle, .reportstores-card, .timeframe');
    if (!clickedInside) {
      this.activePopover = -1;
    }
  }

  constructor(
    public shared: Sharedservice, public setdates: Setdates, private comm: common, private cp: CurrencyPipe, private toast: ToastService
  ) {
    this.shared.setTitle(this.comm.titleName + '-Used Vehicle Stocking');
    this.selectedBlock = [];
    this.selectedBlock.push(this.blocks[1])

    let today = new Date();
    let enddate = new Date(today.setDate(today.getDate() - 1));
    this.month = new Date(enddate)
    this.DuplicatDate = new Date(this.month)
    this.lastyear = new Date()
    this.lastyear.setFullYear(this.DuplicatDate.getFullYear() - 1);
    this.maxDate = new Date();
    this.minDate = new Date();
    this.minDate.setFullYear(this.maxDate.getFullYear() - 3);
    this.maxDate.setMonth(this.maxDate.getMonth());
    this.initializeDates('MTD')

    if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
      this.groupId = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Preferences
      // this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids.split(',')
    }
    if (this.shared.common.groupsandstoresAll.length > 0) {
      this.groupsArray = this.shared.common.groupsandstoresAll.filter((val: any) => val.sg_id != this.shared.common.reconID);
      this.stores = this.shared.common.groupsandstoresAll.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
      this.storeIds = this.stores.map(function (a: any) { return a.ID; })
      this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_Name : this.groupName = ''
      this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
      this.getStoresandGroupsValues()
    }


    this.setHeaderData();
    this.getInventoryData();
  }

  ngOnInit(): void {

  }

  initializeDates(type: any) {
    let dates: any = this.setdates.setDates(type, this.month)
    this.FromDate = dates[0];
    this.ToDate = dates[1];
    localStorage.setItem('time', type);
    this.DateType = type
  }

  getInventoryData() {
    console.log(this.storeIds);

    if (this.storeIds != undefined && this.storeIds.length > 0) {
      this.shared.spinner.show();
      this.GetData();
    } else {
      // this.shared.spinner.hide();
    }
  }

  GetData() {
    this.InventoryData = [];
    this.NoInventoryData = ''
    this.DuplicatDate = new Date(this.month)
    this.lastyear = new Date()
    this.lastyear.setFullYear(this.DuplicatDate.getFullYear() - 1);


    console.log(this.month, this.FromDate, this.lastyear, this.selectedBlock, '...................................');

    this.CIBlock = this.ShowHideCI
    const obj = {
      Stores: this.storeIds.toString(),
      // StartDate: this.shared.datePipe.transform(this.month, 'MM') + '-' + '01' + '-' + this.shared.datePipe.transform(this.month, 'YYYY'),
      // EndDate: this.shared.datePipe.transform(this.month, 'MM') + '-' + lastday + '-' + this.shared.datePipe.transform(this.month, 'YYYY'),
      StartDate: this.shared.datePipe.transform(this.FromDate, 'MM-dd-yyyy'),
      EndDate: this.shared.datePipe.transform(this.ToDate, 'MM-dd-yyyy'),
      TopCount: this.selectedBlock[0].value == 'All' ? 10000 : this.selectedBlock[0].value

    };
    this.shared.api
      .postmethod(this.comm.routeEndpoint + 'GetUsedVehicleStocking', obj)
      .subscribe(
        (res) => {
          this.Updatecount = 0
          const currentTitle = document.title;
          this.InventoryData = [];
          if (res && res.response && res.response.length > 0) {
            this.InventoryData = res.response;
            console.log(res.response, '............');
            this.InventoryData.some(function (x: any) {
              if (x.Data2 != undefined && x.Data2 != '' && x.Data2 != null) {
                x.Data2 = JSON.parse(x.Data2);
                x.Data2 = x.Data2.sort((a: any, b: any) => a.Slno - b.Slno)
              }
              // if (x.YTD != undefined && x.YTD != '' && x.YTD != null) {
              //   x.YTD = JSON.parse(x.YTD);
              // }
            });
            console.log(this.InventoryData);

            this.shared.spinner.hide();
            this.NoInventoryData = '';
          } else {
            this.shared.spinner.hide();
            this.NoInventoryData = 'No Data Found!!';
          }
        },
        (error) => {
          this.toast.show('502 Bad Gate Way Error', 'danger', 'Error');
          this.shared.spinner.hide();
          this.NoInventoryData = 'No Data Found';
        }
      );
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
  popupvalues: any = { Model: '', Year: '', type: '' }
  UsedVehicleStockingDetails: any = []
  UsedVehicleStockingDetailsNoData: any = ''
  aspopup(tmp: any, mainData: any, Year: any, Model: any) {
    // this.popupvalues.Store = mainData.Store_Name;
    this.popupvalues.Year = Year;
    this.popupvalues.Model = Model;
    this.UsedVehicleStockingDetailsNoData = '';
    this.UsedVehicleStockingDetails = []
    this.popupReference = this.shared.ngbmodal.open(tmp, { size: 'xl', backdrop: 'static', keyboard: true, centered: true, modalDialogClass: 'custom-modal' })
    this.GetAcquisitionDetails(mainData.Storeid, mainData.Make, Year, Model)
  }
  closePopup() {
    if (this.popupReference) {
      this.popupReference.close();
    }
  }

  GetAcquisitionDetails(store: any, Make: any, Year: any, Model: any) {
    const obj = {
      "StoreID": this.storeIds.toString(),
      "Year": Year,
      "Make": Model,
      "Model": Model,
      'StartDate': this.shared.datePipe.transform(this.FromDate, 'MM-dd-yyyy'),
      'EndDate': this.shared.datePipe.transform(this.ToDate, 'MM-dd-yyyy'),
    }
    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetUsedVehicleStockingDetails', obj).subscribe(
      (res) => {
        if (res.response && res.response.length > 0) {
          this.UsedVehicleStockingDetails = res.response
          this.UsedVehicleStockingDetailsNoData = ''
        } else {
          this.UsedVehicleStockingDetailsNoData = 'No Data Found'
        }
      },
      (error) => {
        this.UsedVehicleStockingDetailsNoData = 'No Data Found'
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

  expandDataArray: any = []
  expandorcollapse(cat: any) {
    const index = this.expandDataArray.findIndex((i: any) => i == cat);
    if (index >= 0) {
      this.expandDataArray.splice(index, 1);
    } else {
      this.expandDataArray.push(cat);
    }
  }
  ngAfterViewInit() {

    this.shared.api.getStoresAll().subscribe((res: any) => {
      if (this.comm.pageName == 'Used Vehicle Stocking') {
        if (res.obj.storesData != undefined) {
          this.groupsArray = res.obj.storesData;
          this.stores = this.shared.common.groupsandstoresAll.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
          this.storeIds = this.stores.map(function (a: any) { return a.ID; })

          this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_name : this.groupName = ''
          this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
          this.getStoresandGroupsValues();
          this.getInventoryData()
        }
      }
    })
    this.excel = this.shared.api.getExportToExcelAllReports().subscribe((res) => {
      if (this.excel != undefined) {
        if (res.obj.title == 'Used Vehicle Stocking') {
          if (res.obj.state == true) {
            this.exportToExcel();
          }
        }
      }
    });
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res) => {
      if (this.email != undefined) {
        if (res.obj.title == 'Used Vehicle Stocking') {
          if (res.obj.stateEmailPdf == true) {
            // this.sendEmailData(res.obj.Email, res.obj.notes, res.obj.from);
          }
        }
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res) => {
      if (this.print != undefined) {
        if (res.obj.title == 'Used Vehicle Stocking') {
          if (res.obj.statePrint == true) {
            // this.GetPrintData();
          }
        }
      }
    });

    this.Pdf = this.shared.api.getExportToPDFAllReports().subscribe((res) => {
      if (this.Pdf != undefined) {
        console.log('PDF');
        if (res.obj.title == 'Used Vehicle Stocking') {
          console.log('PDF Used Vehicle Stocking');
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

  onOpenCalendar(container: any) {
    container.setViewMode('month');
    container.monthSelectHandler = (event: any): void => {
      container.value = event.date;
      this.month = event.date;
      return;
    };
  }
  changeDate(e: any) {
    console.log(e);
    this.month = e;
  }

  setHeaderData() {
    const headerdata = {
      title: 'Used Vehicle Stocking',
      stores: this.storeIds,
      groups: this.groups,
      month: this.month,
      show: this.selectedBlock
    };
    this.shared.api.SetHeaderData({
      obj: headerdata,
    });

  }
  ngOnDestroy() {
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





  multipleorsingle(block: any, e: any) {
    if (block == 'R') {
      this.Updatecount = 1
      this.selectedBlock = []
      this.selectedBlock.push(e)
    }
    if (block == 'CI') {
      this.ShowHideCI = e;
      this.Updatecount = this.Updatecount + 9
    }

  }


  viewreport() {
    this.activePopover = -1
    if (this.storeIds.length == 0) {
      this.toast.show('Please select atleast one store', 'warning', 'Warning');
    }
    else {
      this.CIBlock = this.ShowHideCI;
      this.DupDateType = this.DateType
      if (this.Updatecount == 9) {
        this.CIBlock = this.ShowHideCI
        this.Updatecount = 0
      } else {
        this.expandDataArray = []
        this.setHeaderData();
        this.getInventoryData()
      }


    }
  }



  exportToExcel() {

  }





}
