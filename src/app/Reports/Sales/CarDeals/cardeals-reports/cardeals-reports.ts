import { Component, HostListener, Input, SimpleChanges } from '@angular/core';

import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { Setdates } from '../../../../Core/Providers/SetDates/setdates';
import { DateRangePicker } from '../../../../CommonFilters/date-range-picker/date-range-picker';
import { Stores } from '../../../../CommonFilters/stores/stores';
@Component({
  selector: 'app-cardeals-reports',
  imports: [SharedModule, DateRangePicker, Stores],
  templateUrl: './cardeals-reports.html',
  styleUrl: './cardeals-reports.scss'
})
export class CardealsReports {

  ngChanges: any = []

  // Report Popup
  @Input() header: any;
  groupId: any = [1];
  storeIds: any = []
  stores: any = []
  groupsArray: any = [];
  storename: any = ''
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  storesFilterData: any = {
    'groupsArray': this.groupsArray, 'groupId': this.groupId, 'storesArray': this.stores, 'storeids': this.storeIds, 'type': 'M', 'others': 'N',
    'groupName': this.groupName, 'storename': this.storename, storecount: null, 'storedisplayname': this.storedisplayname
  };
  Performance: string = 'Load';

  minDate!: Date;
  maxDate!: Date;
  DateType: any = 'MTD';
  displaytime: any = '';
  FromDate: any = '';
  ToDate: any = '';
  Dates: any = {
    'FromDate': this.FromDate, 'ToDate': this.ToDate, "MaxDate": this.maxDate, 'MinDate': this.minDate, 'DateType': this.DateType, 'DisplayTime': this.displaytime,
    Types: [
      { 'code': 'MTD', 'name': 'MTD' },
      { 'code': 'YTD', 'name': 'YTD' },
      { 'code': 'PYTD', 'name': 'PYTD' },
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

  ngOnChanges(changes: SimpleChanges) {
    console.log(changes, 'Report');

    this.ngChanges = changes['header'].currentValue[0];
    this.FromDate = this.ngChanges.fromDate;
    this.ToDate = this.ngChanges.toDate;
    this.neworused = this.ngChanges.dealType;
    this.retailorlease = this.ngChanges.saleType;
    this.dealstatus = this.ngChanges.dealStatus;
    this.setDates(this.ngChanges.datevaluetype)

  }
  constructor(private shared: Sharedservice, private datesSrvc: Setdates) { }

  ngOnInit() {
    this.getGroups();
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
  ngAfterViewInit() {
    // this.shared.api.getStores().subscribe((res: any) => {
    //   if (this.shared.common.pageName == 'Car Deals') {
    //     if (res.obj.storesData != undefined) {
    //       this.groupsArray = res.obj.storesData;
    //       this.groupId = this.ngChanges.groups
    //       // console.log(this.groupId,'....');
    //       this.getStoresandGroupsValues();
    //     }
    //   }
    // })
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.shared.common.pageName == 'Car Deals') {
        if (res.obj.storesData != undefined) {
          this.groupsArray = res.obj.storesData;
          this.groupId = this.ngChanges.groups;
          this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
          this.storeIds = this.ngChanges.storeIds;
          this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_name : this.groupName = ''
          this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
          this.getStoresandGroupsValues()
        }
      }
    })

  }

  
  getGroups() {
    // console.log(this.shared.common.pageName, this.shared.common.groupsandstores);

    if (this.shared.common.groupsandstores != undefined) {
      if (this.shared.common.groupsandstores.length > 0) {
        this.groupsArray = this.shared.common.groupsandstores.filter((val: any) => val.sg_id != this.shared.common.reconID);
        this.groupId = this.ngChanges.groups
        this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
        this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_Name : this.groupName = ''
        this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
        // console.log(this.stores, this.groupsArray, 'Stores and Groups');
        this.getStoresandGroupsValues()
        // this.StoresData(this.ngChanges)
      }
    }
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
  StoresData(data: any) {
    this.storeIds = data.storeids;
    this.groupId = data.groupId;
    this.storename = data.storename;
    this.groupName = data.groupName;
    this.storecount = data.storecount;
    this.storedisplayname = data.storedisplayname;
    console.log(data,'Data.');
    


  }
  setDates(type: any) {
    // localStorage.setItem('time', type);
    // this.datevaluetype=
    console.log(type);
    if (type != 'C') {
      this.displaytime = 'Time Frame (' + this.Dates.Types.filter((val: any) => val.code == type)[0].name + ')';
    }
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
  updatedDates(data: any) {
    console.log(data);
    this.FromDate = data.FromDate;
    this.ToDate = data.ToDate;
    this.DateType = data.DateType;
    this.displaytime = data.DisplayTime
  }
  neworused: any = ['New', 'Used'];
  retailorlease: any = ['Retail ', 'Lease ', 'Misc '];
  includeorexclude: any = [''];
  dealstatus: any = ['Delivered', 'Capped', 'Finalized'];
  selecttarget: any = ['F'];
  toporbottom: any = ['T'];
  includecharges: any = [''];
  Transactorgl: any = ['T'];
  dealTypeData: any = ['All', 'Retail and Lease', 'Retail', 'Lease', 'Wholesale', 'Misc', 'Fleet', 'Demo', 'Special Order', 'Rental', 'Dealer Trade']
  otherblocks: any = ['Retail and Lease']
  QISearchName: any = '';
  Acquisition: any = ['All']
  multipleorsingle(block: any, e: any) {
    if (block == 'NU') {
      const index = this.neworused.findIndex((i: any) => i == e);

      if (index >= 0) {
        // if (e == 'C') {
        //   this.neworused = []
        // }
        this.neworused.splice(index, 1);
      } else {
        if (e == 'C') {
          this.neworused = []
        } else {
          let cindex = this.neworused.findIndex((i: any) => i == 'C')
          this.neworused.findIndex((i: any) => i == 'C') >= 0 ? this.neworused.splice(cindex, 1) : ''
        }
        this.neworused.push(e);
      }
    }
    if (block == 'RL') {
      const index = this.retailorlease.findIndex((i: any) => i == e);
      if (index >= 0) {
        if (e == 'All') {
          // this.retailorlease.splice(index, 1);
          this.retailorlease = []
          // alert('Please select atleast one Deal Type');

        } else {
          this.retailorlease.splice(index, 1);
          let allindex = this.retailorlease.findIndex((i: any) => i == 'All');
          if (allindex >= 0) {
            this.retailorlease.splice(allindex, 1);
          }
        }
      } else {
        this.otherblocks = ['Retail and Lease']
        if (e == 'All') {
          const dealdata = JSON.stringify(this.dealTypeData)
          this.retailorlease = JSON.parse(dealdata)
        } else {
          this.retailorlease.push(e);
          // if (this.retailorlease.length == 0) {
          //   alert('Please select atleast one Deal Type');

          // }
        }
      }
    }
    if (block == 'PH') {
      this.includeorexclude = [];
      this.includeorexclude.push(e);
    }
    if (block == 'OB') {
      this.otherblocks = []
      this.otherblocks.push(e)
      if (e == 'Retail and Lease') {
        this.retailorlease = ['Retail', 'Lease', 'Misc', 'Special Order'];
      } else {
        this.retailorlease = []
      }
      // this.QISearchName = ''
      // this.setHeaderReportData()
    }
    if (block == 'DS') {
      const index = this.dealstatus.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.dealstatus.splice(index, 1);
      } else {
        this.dealstatus.push(e);
      }
    }
    if (block == 'ST') {
      const index = this.selecttarget.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.selecttarget.splice(index, 1);
      } else {
        this.selecttarget.push(e);
      }
    }
    if (block == 'TB') {
      this.toporbottom = [];
      this.toporbottom.push(e);
    }
    if (block == 'IC') {
      this.includecharges = [];
      this.includecharges.push(e);
    }
    if (block == 'SRC') {
      this.Transactorgl = [];
      this.Transactorgl.push(e);
    }

  }
  viewreport() {
    this.activePopover = -1

    if (this.retailorlease.length == 0) {
      alert('Please select atleast one Deal Type');
    }
    else if (this.dealstatus.length == 0) {
      alert('Please select atleast one Deal Status');
    }
    else {
      const data = {
        Reference: 'Car Deals',
        FromDate: this.shared.datePipe.transform(this.FromDate, 'MM-dd-yyyy'),
        ToDate: this.shared.datePipe.transform(this.ToDate, 'MM-dd-yyyy'),
        // TotalReport: this.toporbottom[0],
        storeValues: this.storeIds.toString(),
        dateType: this.DateType,
        dealType: this.neworused,
        saleType: this.retailorlease,
        dealStatus: this.dealstatus,

        groups: this.groupId,
        otherblock: this.otherblocks,
        search: this.QISearchName
        // gv: this.gridview,
        // acquisition: this.Acquisition.length == this.acquisitionsource.length ? aqusrc : this.Acquisition,
        // groups: this.selectedGroups.toString(),
        // otherstoreids: this.selectedotherstoreids
      };
      console.log(data);
      this.shared.api.SetReports({
        obj: data,
      });
      this.close();
    }

  }

  close() {
    this.Performance = 'Unload';
  }
}

