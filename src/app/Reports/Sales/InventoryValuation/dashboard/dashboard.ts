import { Component, OnInit, ViewChild, ElementRef, HostListener } from '@angular/core';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { Setdates } from '../../../../Core/Providers/SetDates/setdates';
import { BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { common } from '../../../../common';
import { Stores } from '../../../../CommonFilters/stores/stores';
import { DateRangePicker } from '../../../../CommonFilters/date-range-picker/date-range-picker';
import { CurrencyPipe } from '@angular/common';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
// import { SalesgrossDetailsComponent } from '../../Sales/Gross/salesgross-details/salesgross-details.component';
const EXCEL_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
const EXCEL_EXTENSION = '.xlsx';

@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [SharedModule, BsDatepickerModule, Stores],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
})
export class Dashboard {
  InventoryValuation: any = [];
  InventoryValuationRaw: any = [];
  FromDate: any = '';
  ToDate: any = '';
  minDate!: Date;
  maxDate!: Date;
  DateType: any = 'MTD';
  displaytime: any = '';
  NoData: any = '';

  //  storeIds: any = []
  stores: any = []
  groupsArray: any = [];
  storename: any = ''
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  groupId: any = 0;

  storesFilterData: any = {
    'groupsArray': this.groupsArray, 'groupId': this.groupId, 'storesArray': this.stores, 'storeids': '1', 'type': 'M', 'others': 'N',
    'groupName': this.groupName, 'storename': this.storename, storecount: null, 'storedisplayname': this.storedisplayname
  };

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


  // filters
  storeIds!: any;
  dateType: any = 'MTD';
  groups: any = 1;
  constructor(public shared: Sharedservice, public setdates: Setdates, private comm: common, private cp: CurrencyPipe,private toast: ToastService,) {
    this.shared.setTitle('Used Write Down');
    this.initializeDates(this.DateType)


    if (typeof window !== 'undefined') {
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


      this.shared.setTitle(this.shared.common.titleName + '-Used Write Down');
      const data = {
        title: 'Used Write Down',
        stores: this.storeIds,
      };
      this.shared.api.SetHeaderData({
        obj: data,
      });

      this.GetData();
      this.setDates(this.DateType)
    }
  }

  ngOnInit(): void { }

  getTotal(frontgross: any[], colname: string) {
    return frontgross.reduce((total, x) => {
      const raw = x?.[colname];
      // Skip null/undefined/empty strings
      if (raw === null || raw === undefined || raw === '') return total;
      // Parse as number (handles strings like "123" or "123.45")
      const n = Number(raw);
      // Skip NaN
      if (Number.isNaN(n)) return total;
      // Add the value (use Math.trunc if you only want integer part)
      return total + n;
    }, 0);
  }

  StoresData(data: any) {
    this.storeIds = data.storeids;
    this.groupId = data.groupId;
    this.storename = data.storename;
    this.groupName = data.groupName;
    this.storecount = data.storecount;
    this.storedisplayname = data.storedisplayname;
  }

  initializeDates(type: any) {
    let dates: any = this.setdates.setDates(type)
    this.FromDate = dates[0];
    this.ToDate = dates[1];
    localStorage.setItem('time', type);
  }

  GetData() {

    this.InventoryValuation = [];
    this.shared.spinner.show();
    const obj = {
      "Stores": this.storeIds.toString(),
    
    };
    const curl = this.shared.getEnviUrl() + 'GetInventoryUsedWritedown';
    this.shared.api.postmethod(this.shared.common.routeEndpoint + 'GetInventoryUsedWritedown', obj).subscribe((res) => {
      if (res.status == 200) {
        if (res.response != undefined) {
          if (res.response.length > 0) {
            this.InventoryValuation = res.response;
            this.shared.spinner.hide();
            this.NoData = '';
            console.log(this.InventoryValuation, 'Used Write Down Data.........');

          } else {
            this.shared.spinner.hide();
            this.NoData = 'No Data Found!!';
          }
        } else {
          this.shared.spinner.hide();
          this.NoData = 'No Data Found!!';
        }
      } else {
        this.shared.spinner.hide();
        this.NoData = 'No Data Found!!';
      }
    },
      (error) => {
        //  this.shared.toaster.error('502 Bad Gate Way Error', '');
        this.shared.spinner.hide();
        this.NoData = 'No Data Found!!';
      }
    );
    // this.InventoryValuationRaw = [
    //   {
    //     id: 1,
    //     stock: "1HU7865",
    //     status: "1 - In Stock",
    //     year: 2020,
    //     make: "NISSAN",
    //     model: "ALTIMA",
    //     age: 61,
    //     acqCost: 30000,
    //     recon: 1800,
    //     glCost: 31800,
    //     netOfPack: 31000,
    //     bookValue: 30200,
    //     potentialWD: -800,
    //     store: "Cadillac of Beverly Hills",
    //     vin: "1N4BL4CVXLC134313",
    //     trim: "2.5 SR",
    //     mileage: 67425,
    //     color: "BLUE"
    //   },
    //   {
    //     id: 2,
    //     stock: "1H3256A",
    //     status: "1 - In Stock",
    //     year: 2019,
    //     make: "NISSAN",
    //     model: "FRONTIER",
    //     age: 65,
    //     acqCost: 40000,
    //     recon: 1200,
    //     glCost: 41200,
    //     netOfPack: 40500,
    //     bookValue: 40000,
    //     potentialWD: -500,
    //     store: "Santa Monica GMC Inc",
    //     vin: "1N6AD0EV0KN743029", // best guess from image; please confirm
    //     trim: "SV",
    //     mileage: 11318,
    //     color: null
    //   },
    //   {
    //     id: 1,
    //     stock: "1HR2412A",
    //     status: "1 - In Stock",
    //     year: 2016,
    //     make: "NISSAN",
    //     model: "PATHFINDER",
    //     age: 69,
    //     acqCost: 20000,
    //     recon: 750,
    //     glCost: 20750,
    //     netOfPack: 20000,
    //     bookValue: 20000,
    //     potentialWD: 0,
    //     store: "Cadillac of Beverly Hills",
    //     vin: "5N1AR2MMXGC65430", // likely missing a digit (VIN should be 17 chars)
    //     trim: "SL",
    //     mileage: 113091,
    //     color: null
    //   },
    //   {
    //     id: 2,
    //     stock: "1HU7837",
    //     status: "1 - In Stock",
    //     year: 2015,
    //     make: "NISSAN",
    //     model: "LEAF",
    //     age: 75,
    //     acqCost: 35000,
    //     recon: 2200,
    //     glCost: 37200,
    //     netOfPack: 36500,
    //     bookValue: 37000,
    //     potentialWD: 0,
    //     store: "Santa Monica GMC Inc",
    //     vin: "1N4AZ0CP9FC301737",
    //     trim: "SV",
    //     mileage: 39893,
    //     color: "BLACK"
    //   },
    //   {
    //     id: 1,
    //     stock: "1H3049B",
    //     status: "1 - In Stock",
    //     year: 2014,
    //     make: "NISSAN",
    //     model: "ROGUE",
    //     age: 84,
    //     acqCost: 20000,
    //     recon: 1450,
    //     glCost: 21450,
    //     netOfPack: 21000,
    //     bookValue: 18000,
    //     potentialWD: -3000,
    //     store: "Cadillac of Beverly Hills",
    //     vin: "5N1AT2MK6EC809460",
    //     trim: "",
    //     mileage: 137569,
    //     color: null
    //   },
    //    {
    //     id: 1,
    //     stock: "1HU7865",
    //     status: "1 - In Stock",
    //     year: 2020,
    //     make: "NISSAN",
    //     model: "ALTIMA",
    //     age: 61,
    //     acqCost: 30000,
    //     recon: 1800,
    //     glCost: 31800,
    //     netOfPack: 31000,
    //     bookValue: 30200,
    //     potentialWD: -800,
    //     store: "Cadillac of Beverly Hills",
    //     vin: "1N4BL4CVXLC134313",
    //     trim: "2.5 SR",
    //     mileage: 67425,
    //     color: "BLUE"
    //   },
    //   {
    //     id: 2,
    //     stock: "1H3256A",
    //     status: "1 - In Stock",
    //     year: 2019,
    //     make: "NISSAN",
    //     model: "FRONTIER",
    //     age: 65,
    //     acqCost: 40000,
    //     recon: 1200,
    //     glCost: 41200,
    //     netOfPack: 40500,
    //     bookValue: 40000,
    //     potentialWD: -500,
    //     store: "Santa Monica GMC Inc",
    //     vin: "1N6AD0EV0KN743029", // best guess from image; please confirm
    //     trim: "SV",
    //     mileage: 11318,
    //     color: null
    //   },
    //   {
    //     id: 1,
    //     stock: "1HR2412A",
    //     status: "1 - In Stock",
    //     year: 2016,
    //     make: "NISSAN",
    //     model: "PATHFINDER",
    //     age: 69,
    //     acqCost: 20000,
    //     recon: 750,
    //     glCost: 20750,
    //     netOfPack: 20000,
    //     bookValue: 20000,
    //     potentialWD: 0,
    //     store: "Cadillac of Beverly Hills",
    //     vin: "5N1AR2MMXGC65430", // likely missing a digit (VIN should be 17 chars)
    //     trim: "SL",
    //     mileage: 113091,
    //     color: null
    //   },
    //   {
    //     id: 2,
    //     stock: "1HU7837",
    //     status: "1 - In Stock",
    //     year: 2015,
    //     make: "NISSAN",
    //     model: "LEAF",
    //     age: 75,
    //     acqCost: 35000,
    //     recon: 2200,
    //     glCost: 37200,
    //     netOfPack: 36500,
    //     bookValue: 37000,
    //     potentialWD: 0,
    //     store: "Santa Monica GMC Inc",
    //     vin: "1N4AZ0CP9FC301737",
    //     trim: "SV",
    //     mileage: 39893,
    //     color: "BLACK"
    //   },
    //   {
    //     id: 1,
    //     stock: "1H3049B",
    //     status: "1 - In Stock",
    //     year: 2014,
    //     make: "NISSAN",
    //     model: "ROGUE",
    //     age: 84,
    //     acqCost: 20000,
    //     recon: 1450,
    //     glCost: 21450,
    //     netOfPack: 21000,
    //     bookValue: 18000,
    //     potentialWD: -3000,
    //     store: "Cadillac of Beverly Hills",
    //     vin: "5N1AT2MK6EC809460",
    //     trim: "",
    //     mileage: 137569,
    //     color: null
    //   }
    // ];

    // this.InventoryValuation = [
    //   {
    //     id: 1,
    //     stock: "1HU7865",
    //     status: "1 - In Stock",
    //     year: 2020,
    //     make: "NISSAN",
    //     model: "ALTIMA",
    //     age: 61,
    //     acqCost: 30000,
    //     recon: 1800,
    //     glCost: 31800,
    //     netOfPack: 31000,
    //     bookValue: 30200,
    //     potentialWD: -800,
    //     store: "Cadillac of Beverly Hills",
    //     vin: "1N4BL4CVXLC134313",
    //     trim: "2.5 SR",
    //     mileage: 67425,
    //     color: "BLUE"
    //   },
    //   {
    //     id: 2,
    //     stock: "1H3256A",
    //     status: "1 - In Stock",
    //     year: 2019,
    //     make: "NISSAN",
    //     model: "FRONTIER",
    //     age: 65,
    //     acqCost: 40000,
    //     recon: 1200,
    //     glCost: 41200,
    //     netOfPack: 40500,
    //     bookValue: 40000,
    //     potentialWD: -500,
    //     store: "Santa Monica GMC Inc",
    //     vin: "1N6AD0EV0KN743029", // best guess from image; please confirm
    //     trim: "SV",
    //     mileage: 11318,
    //     color: null
    //   },
    //   {
    //     id: 1,
    //     stock: "1HR2412A",
    //     status: "1 - In Stock",
    //     year: 2016,
    //     make: "NISSAN",
    //     model: "PATHFINDER",
    //     age: 69,
    //     acqCost: 20000,
    //     recon: 750,
    //     glCost: 20750,
    //     netOfPack: 20000,
    //     bookValue: 20000,
    //     potentialWD: 0,
    //     store: "Cadillac of Beverly Hills",
    //     vin: "5N1AR2MMXGC65430", // likely missing a digit (VIN should be 17 chars)
    //     trim: "SL",
    //     mileage: 113091,
    //     color: null
    //   },
    //   {
    //     id: 2,
    //     stock: "1HU7837",
    //     status: "1 - In Stock",
    //     year: 2015,
    //     make: "NISSAN",
    //     model: "LEAF",
    //     age: 75,
    //     acqCost: 35000,
    //     recon: 2200,
    //     glCost: 37200,
    //     netOfPack: 36500,
    //     bookValue: 37000,
    //     potentialWD: 0,
    //     store: "Santa Monica GMC Inc",
    //     vin: "1N4AZ0CP9FC301737",
    //     trim: "SV",
    //     mileage: 39893,
    //     color: "BLACK"
    //   },
    //   {
    //     id: 1,
    //     stock: "1H3049B",
    //     status: "1 - In Stock",
    //     year: 2014,
    //     make: "NISSAN",
    //     model: "ROGUE",
    //     age: 84,
    //     acqCost: 20000,
    //     recon: 1450,
    //     glCost: 21450,
    //     netOfPack: 21000,
    //     bookValue: 18000,
    //     potentialWD: -3000,
    //     store: "Cadillac of Beverly Hills",
    //     vin: "5N1AT2MK6EC809460",
    //     trim: "",
    //     mileage: 137569,
    //     color: null
    //   },
    //    {
    //     id: 1,
    //     stock: "1HU7865",
    //     status: "1 - In Stock",
    //     year: 2020,
    //     make: "NISSAN",
    //     model: "ALTIMA",
    //     age: 61,
    //     acqCost: 30000,
    //     recon: 1800,
    //     glCost: 31800,
    //     netOfPack: 31000,
    //     bookValue: 30200,
    //     potentialWD: -800,
    //     store: "Cadillac of Beverly Hills",
    //     vin: "1N4BL4CVXLC134313",
    //     trim: "2.5 SR",
    //     mileage: 67425,
    //     color: "BLUE"
    //   },
    //   {
    //     id: 2,
    //     stock: "1H3256A",
    //     status: "1 - In Stock",
    //     year: 2019,
    //     make: "NISSAN",
    //     model: "FRONTIER",
    //     age: 65,
    //     acqCost: 40000,
    //     recon: 1200,
    //     glCost: 41200,
    //     netOfPack: 40500,
    //     bookValue: 40000,
    //     potentialWD: -500,
    //     store: "Santa Monica GMC Inc",
    //     vin: "1N6AD0EV0KN743029", // best guess from image; please confirm
    //     trim: "SV",
    //     mileage: 11318,
    //     color: null
    //   },
    //   {
    //     id: 1,
    //     stock: "1HR2412A",
    //     status: "1 - In Stock",
    //     year: 2016,
    //     make: "NISSAN",
    //     model: "PATHFINDER",
    //     age: 69,
    //     acqCost: 20000,
    //     recon: 750,
    //     glCost: 20750,
    //     netOfPack: 20000,
    //     bookValue: 20000,
    //     potentialWD: 0,
    //     store: "Cadillac of Beverly Hills",
    //     vin: "5N1AR2MMXGC65430", // likely missing a digit (VIN should be 17 chars)
    //     trim: "SL",
    //     mileage: 113091,
    //     color: null
    //   },
    //   {
    //     id: 2,
    //     stock: "1HU7837",
    //     status: "1 - In Stock",
    //     year: 2015,
    //     make: "NISSAN",
    //     model: "LEAF",
    //     age: 75,
    //     acqCost: 35000,
    //     recon: 2200,
    //     glCost: 37200,
    //     netOfPack: 36500,
    //     bookValue: 37000,
    //     potentialWD: 0,
    //     store: "Santa Monica GMC Inc",
    //     vin: "1N4AZ0CP9FC301737",
    //     trim: "SV",
    //     mileage: 39893,
    //     color: "BLACK"
    //   },
    //   {
    //     id: 1,
    //     stock: "1H3049B",
    //     status: "1 - In Stock",
    //     year: 2014,
    //     make: "NISSAN",
    //     model: "ROGUE",
    //     age: 84,
    //     acqCost: 20000,
    //     recon: 1450,
    //     glCost: 21450,
    //     netOfPack: 21000,
    //     bookValue: 18000,
    //     potentialWD: -3000,
    //     store: "Cadillac of Beverly Hills",
    //     vin: "5N1AT2MK6EC809460",
    //     trim: "",
    //     mileage: 137569,
    //     color: null
    //   }
    // ];

  }



  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    }
    return false;
  }

  isDesc: boolean = false;
  column: string = 'CategoryName';
  sort(property: any) {
    this.isDesc = !this.isDesc; //change the direction
    this.column = property;
    let direction = this.isDesc ? 1 : -1;

    this.InventoryValuation.sort((a: any, b: any) => {
      const valA = a[property] ?? ''; // replace null/undefined with empty string
      const valB = b[property] ?? '';

      if (valA < valB) {
        return -1 * direction;
      } else if (valA > valB) {
        return 1 * direction;
      } else {
        return 0;
      }
    });

  }
  SPRstate: any;
  ngAfterViewInit(): void {
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.comm.pageName == 'Used Write Down') {
        if (res.obj.storesData != undefined) {
          this.groupsArray = res.obj.storesData;
          this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
          this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_name : this.groupName = ''
          this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
          this.getStoresandGroupsValues()
        }
      }
    })

    this.shared.api.getExportToExcelAllReports().subscribe((res) => {
      this.SPRstate = res.obj.state;
      if (res.obj.title == 'Used Write Down') {
        if (res.obj.state == true) {
          // this.exportToExcel();
        }
      }
    });

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



  openCustomPicker(datepicker: any) {
    this.DateType = 'C';
    localStorage.setItem('time', this.DateType);
    datepicker.toggle();
  }


  dateRangeCreated($event: any) {
    if ($event) {
      this.FromDate = this.shared.datePipe.transform($event[0], 'MM-dd-yyyy');
      this.ToDate = this.shared.datePipe.transform($event[1], 'MM-dd-yyyy');
      this.bsRangeValue = [this.FromDate, this.ToDate];
      if (this.DateType === 'C') this.custom = true;
    }
  }

  updatedDates(data: any) {
    // console.log(data);
    this.FromDate = data.FromDate;
    this.ToDate = data.ToDate;
    this.DateType = data.DateType;
    this.displaytime = data.DisplayTime
  }




  // comments code

  Scrollpercent: any = 0;
  scrollCurrentposition: any = 0;
  @ViewChild('scrollcent') scrollcent!: ElementRef;
  updateVerticalScroll(event: any): void {
    this.scrollCurrentposition = event.target.scrollTop;
    const scrollDemo = document.querySelector('#scrollcent') as HTMLElement;
    this.Scrollpercent = Math.round(
      (event.target.scrollTop / (event.target.scrollHeight - scrollDemo.clientHeight)) * 100
    );
  }






  //////////////////////  REPORT CODE /////////////////////////////////////////////
  activePopover: number = -1;
  custom: boolean = false;
  bsRangeValue!: Date[];
  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent) {
    const clickedInside = (event.target as HTMLElement).closest('.dropdown-toggle, .reportstores-card, .timeframe');
    if (!clickedInside) {
      this.activePopover = -1;
    }
  }

  togglePopover(popoverIndex: number) {
    this.activePopover = this.activePopover === popoverIndex ? -1 : popoverIndex;
  }


  setDates(type: any) {
    this.displaytime = 'Time Frame (' + this.Dates.Types.filter((val: any) => val.code == type)[0].name + ')';
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
  close() {
    this.shared.ngbmodal.dismissAll();
  }

  // Final Apply
  viewreport() {
    this.activePopover = -1;

    if (this.storeIds && this.storeIds.length > 0) {
      const data = {
        Reference: 'Used Write Down',
        FromDate: this.FromDate,
        ToDate: this.ToDate,
        storeValues: this.storeIds.toString(),
        dateType: this.DateType,
        groups: this.groupId.toString(),

      };
      this.shared.api.SetReports({
        obj: data,
      });
      this.close();
      this.GetData();
    //   this.InventoryValuation= this.InventoryValuationRaw.filter((item: any) =>
    //   this.storeIds.includes(item.id)
    // );
    } else {

      this.toast.show('Please Select Atleast One Store', 'warning', 'Warning');
    }




  }

}
