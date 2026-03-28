import { Component, OnInit, ViewChild, ElementRef, HostListener } from '@angular/core';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { Setdates } from '../../../../Core/Providers/SetDates/setdates';
import { BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { common } from '../../../../common';
import { DateRangePicker } from '../../../../CommonFilters/date-range-picker/date-range-picker';
import { Subscription } from 'rxjs';
import { FilterPipe } from '../../../../Core/Providers/filterpipe/filter.pipe';
import { Stores } from '../../../../CommonFilters/stores/stores';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [SharedModule, BsDatepickerModule, DateRangePicker, FilterPipe,Stores],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss'
})
export class Dashboard {
  Appointmentdata: any = [];
  NoData: any = false;

  appointment: any = [];
  selectedStore = '0';
  selectedstrid = 0;
  // stores: any = [];
  PageCount = 1;
  LastCount!: boolean;
  Pagination: boolean = false;
  storeIds: any = '0';
  // lenders edit

  // groups: any = 1
  filteredCustomers: any;
  Appointmentsearch: any;
  searcdata: any;

  FromDate: any = '';
  ToDate: any = '';
  todaytitle: any;

  Pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;

  minDate!: Date;
  maxDate!: Date;
  DateType: any = 'C';
  displaytime: any = new Date();
  Dates: any = {
    'FromDate': this.FromDate, 'ToDate': this.ToDate, "MaxDate": this.maxDate, 'MinDate': this.minDate, 'DateType': this.DateType, 'DisplayTime': this.displaytime,
    Types: [
      // { 'code': '', 'name': '' },
      // { 'code': 'QTD', 'name': 'QTD' },
      // { 'code': 'YTD', 'name': 'YTD' },
      // { 'code': 'PYTD', 'name': 'PYTD' },
      // { 'code': 'LY', 'name': 'Last Year' },
      // { 'code': 'LM', 'name': 'Last Month' },
      // { 'code': 'PM', 'name': 'Same Month PY' },
    ]
  }
  bsRangeValue!: Date[];
  StoreValues: any = '0';
  popup: any = [{ type: 'Popup' }];
  groups: any = 1;
  gridvisibility: any;
  // bsRangeValue!: Date[];
  groupsArray: any = [];
  storename: any = ''
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  groupId: any = 0;
  stores: any = [];
  			  otherStoresArray: any = [];
  otherStoreIds: any = [];
  storesFilterData: any = {
    'groupsArray': this.groupsArray, 'groupId': this.groupId, 'storesArray': this.stores, 'storeids': '1', 'type': 'M', 'others': 'Y',  otherStoresArray: this.otherStoresArray, otherStoreIds: this.otherStoreIds,
    'groupName': this.groupName, 'storename': this.storename, storecount: null, 'storedisplayname': this.storedisplayname
  };
  constructor(
    public shared: Sharedservice, public setdates: Setdates, private comm: common,private toast: ToastService,
  ) {
    if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
      // this.StoreValues = '1,2';
      this.groupId = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Preferences
      this.StoreValues = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids.split(',')

        this.otherStoreIds = JSON.parse(localStorage.getItem('otherstoreids')!);
      if (this.StoreValues.toString().indexOf(',') > 0) {
        this.gridvisibility = 'DL';
      } else {
        this.gridvisibility = 'SL';
      }
    }
    this.titleSetting()
    this.shared.setTitle(this.shared.common.titleName + '- Service Appointments');
    this.todaytitle = this.shared.datePipe.transform(new Date(), 'MM/dd/yyyy');
    let today = new Date();
    const tomorrow = new Date();
    tomorrow.setDate(today.getDate());
    this.FromDate = this.shared.datePipe.transform(today, 'MM/dd/yyyy');
    this.ToDate = this.shared.datePipe.transform(tomorrow, 'MM/dd/yyyy');
    this.setDates(this.DateType)
    // if (localStorage.getItem('Fav') != 'Y') {
    this.setHeaderData();
    this.GetAppointment();
    // }
  }
  dynamicTitle = 'Service Appointments'
  titleSetting() {
    if (this.DateType == 'C') {
      if (this.todaytitle == this.shared.datePipe.transform(this.FromDate, 'MM/dd/yyyy') && this.todaytitle == this.shared.datePipe.transform(this.ToDate, 'MM/dd/yyyy')) {
        this.dynamicTitle = 'Service Appointments'
      }
      else {
        this.dynamicTitle = 'Service Appointments'
      }
    }

  this.setHeaderData()
  }

  setHeaderData() {
    const data = {
      title: this.dynamicTitle,
      stores: this.storeIds,
      groups: this.groupId,
      fromdate: this.FromDate,
      todate: this.ToDate,
      otherstoreids: this.otherStoreIds

    };
    this.shared.api.SetHeaderData({
      obj: data,
    });
  }
  ngOnInit(): void { }


  searchQuery: string = '';
  isDesc: boolean = false;
  column: string = '';

  sort(property: any) {
    this.isDesc = !this.isDesc; //change the direction
    this.column = property;
    let direction = this.isDesc ? 1 : -1;
    // // console.log(property)
    this.Appointmentdata.sort(function (a: any, b: any) {
      if (a[property] < b[property]) {
        return -1 * direction;
      } else if (a[property] > b[property]) {
        return 1 * direction;
      } else {
        return 0;
      }
    });
  }


  GetAppointment() {
    this.Appointmentdata = [];
    this.NoData = false;
    this.shared.spinner.show();
    this.FromDate = this.FromDate.replace(/-/g, '/');
    this.ToDate = this.ToDate.replace(/-/g, '/');
    let obj = {};
    if (localStorage.getItem('DetailsObject') != null) {
      // console.log('locstg', JSON.parse(localStorage.getItem('DetailsObject')!));
      const InvObj = JSON.parse(localStorage.getItem('DetailsObject')!);
      obj = {
        DealerId: InvObj.dataobj.Data1,
        StartDate: InvObj.dataobj.Data2,
        EndDate: InvObj.dataobj.Data2,
      };
      this.FromDate = InvObj.dataobj.Data2
      this.ToDate = InvObj.dataobj.Data2
      this.storeIds =  InvObj.dataobj.Data1
      if (this.todaytitle == this.shared.datePipe.transform(this.FromDate, 'MM/dd/yyyy') && this.todaytitle == this.shared.datePipe.transform(this.ToDate, 'MM/dd/yyyy')) {
        this.dynamicTitle = 'Service Appointments'
      }
      else {
        this.dynamicTitle = 'Service Appointments'
      }
      this.setHeaderData()
    } else {
      obj = {
        DealerId: [...this.StoreValues, ...this.otherStoreIds],
        StartDate: this.FromDate,
        EndDate: this.ToDate,
      };
      console.log(obj);

    }
    const curl = this.shared.getEnviUrl() + this.comm.routeEndpoint + 'GetServiceAppointments';
    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetServiceAppointments', obj).subscribe(
      (res: { message: any; status: number; response: string | any[] | undefined; }) => {
        const currentTitle = document.title;
        this.shared.api.logSaving(curl, {}, '', res.message, currentTitle);
        if (res.status == 200) {
          localStorage.removeItem('DetailsObject');
          if (res.response != undefined) {
            if (res.response.length > 0) {
              this.shared.spinner.hide();
              this.Appointmentdata = res.response;
              this.Appointmentsearch = res.response;
              console.log(this.Appointmentdata, 'Appointment');
              this.appointment = res.response;
              if (this.appointment.length == 0) {
                this.Pagination = false;
                this.LastCount = false;
                this.NoData = true;
              } else {
                this.Pagination = true;
                this.LastCount = true;
              }
            }
            else {
              this.shared.spinner.hide();
              this.NoData = true;
            }
          }
          else {
            // this.toast.error('Empty Response');
            this.NoData = true;
            this.shared.spinner.hide();
          }

        } else {
          // this.toast.error(res.status, '');
          this.shared.spinner.hide();
          this.NoData = true;
        }
      },
      (error: any) => {
        // this.toast.error('502 Bad Gate Way Error', '');
        this.shared.spinner.hide();
        this.NoData = true;
      }
    );
  }



  ngAfterViewInit() {

    this.shared.api.getStores().subscribe((res: any) => {
      if (this.shared.common.pageName == this.dynamicTitle) {
        if (res.obj.storesData != undefined) {
          this.groupsArray = res.obj.storesData;
          this.otherStoresArray = this.shared.common.OtherStoresData[0].Stores

          this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
          this.StoreValues.length == this.stores.length ? this.groupName = this.stores[0].sg_name : this.groupName = ''
          this.StoreValues.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.StoreValues)[0].storename : this.storename = ''
          this.getStoresandGroupsValues()
        }
      }
    })
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res: { obj: { title: string; stateEmailPdf: boolean; }; }) => {
      if (this.email != undefined) {
        if (res.obj.title == this.dynamicTitle) {
          if (res.obj.stateEmailPdf == true) {
            // this.sendEmailData(res.obj.Email, res.obj.notes,res.obj.from);
          }
        }
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res: { obj: { title: string; statePrint: boolean; }; }) => {
      if (this.print != undefined) {
        if (res.obj.title == this.dynamicTitle) {
          if (res.obj.statePrint == true) {
            // this.GetPrintData();
          }
        }
      }
    });
    this.Pdf = this.shared.api.getExportToPDFAllReports().subscribe((res: { obj: { title: string; statePDF: boolean; }; }) => {
      if (this.Pdf != undefined) {

        if (res.obj.title == this.dynamicTitle) {
          if (res.obj.statePDF == true) {
            // this.generatePDF();
          }
        }
      }
    });

    this.excel = this.shared.api.getExportToExcelAllReports().subscribe((res: { obj: { state: boolean; title: string; }; }) => {
      if (this.excel != undefined) {
        if (res.obj.title == this.dynamicTitle) {
          if (res.obj.state == true) {
            this.exportToExcel();
          }
        }
      }
    });
  }
  getStoresandGroupsValues() {
    this.storesFilterData.groupsArray = this.groupsArray;
    this.storesFilterData.groupId = this.groupId;
    this.storesFilterData.storesArray = this.stores;
    this.storesFilterData.storeids = this.StoreValues;
    this.storesFilterData.groupName = this.groupName;
    this.storesFilterData.storename = this.storename;
    this.storesFilterData.storecount = this.storecount;
    this.storesFilterData.storedisplayname = this.storedisplayname;
this.storesFilterData.otherStoreIds = this.otherStoreIds;
    this.storesFilterData.otherStoresArray = this.otherStoresArray;
    this.storesFilterData = {
      groupsArray: this.groupsArray,
      groupId: this.groupId,
      storesArray: this.stores,
      storeids: this.StoreValues,
      groupName: this.groupName,
      storename: this.storename,
      storecount: this.storecount,
      storedisplayname: this.storedisplayname,
      'type': 'M', 'others': 'Y', otherStoresArray: this.otherStoresArray,
      otherStoreIds: this.otherStoreIds,
    };

    // this.setHeaderData();
    // this.GetData();

  }
  StoresData(data: any) {
    this.StoreValues = data.storeids;
    this.groupId = data.groupId;
    this.storename = data.storename;
    this.groupName = data.groupName;
    this.storecount = data.storecount;
    this.otherStoreIds = data.otherStoreIds;
    this.storedisplayname = data.storedisplayname;


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
  Scrollpercent: any = 0;
  scrollCurrentposition: any = 0
  @ViewChild('scrollcent') scrollcent!: ElementRef;

  updateVerticalScroll(event: any): void {

    this.scrollCurrentposition = event.target.scrollTop
    const scrollDemo = document.querySelector('#scrollcent') as HTMLElement;
    this.Scrollpercent = Math.round(
      (event.target.scrollTop /
        (event.target.scrollHeight - scrollDemo.clientHeight)) *
      100
    );
  }


  //------Reports-----------//
  activePopover: number = -1;
  Teams = 'S';
  selectedstorevalues: any = [];
  labortypes: any = [];
  selectedlabortypevalues: any = [];
  Performance: any = 'Load';
  changesvalues: any = [];
  AllLabortypes: boolean = true;
  toporbottom: any = ['T'];

  initializeDates(type: any) {
    let dates: any = this.setdates.setDates(type)
    this.FromDate = dates[0];
    this.ToDate = dates[1];
    localStorage.setItem('time', type);
    this.setDates(type)
  }



  updatedDates(data: any) {
    console.log(data);
    this.FromDate = data.FromDate;
    this.ToDate = data.ToDate;
    this.DateType = data.DateType;
    this.displaytime = data.DisplayTime
    this.bsRangeValue = [new Date(this.FromDate), new Date(this.ToDate)];

  }
  open() {
    this.Dates = { ...this.Dates }
    this.setDates(this.DateType)
  }
  setDates(type: any) {
    this.DateType == 'C' ? this.displaytime = ' (  ' + this.shared.datePipe.transform(this.FromDate, 'MM/dd/yyyy') + ' - ' + this.shared.datePipe.transform(this.ToDate, 'MM/dd/yyyy') + ' ) ':
      this.displaytime = 'Time Frame (' + this.Dates.Types.filter((val: any) => val.code == type)[0].name + ')';
    // this.maxDate = new Date();
    // this.minDate = new Date();
    // this.minDate.setFullYear(this.maxDate.getFullYear() - 3);
    // this.maxDate.setDate(this.maxDate.getDate());
    this.Dates.FromDate = this.FromDate;
    this.Dates.ToDate = this.ToDate;
    this.Dates.MinDate = this.minDate;
    this.Dates.MaxDate = this.maxDate;
    this.Dates.DateType = this.DateType;
    this.Dates.DisplayTime = this.displaytime;
    console.log(this.FromDate, this.ToDate, this.DateType, this.displaytime, '..............');
  }
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

  viewreport() {
    if (!this.StoreValues || this.StoreValues.length === 0) {
    
      this.toast.show('Please select atleast one store', 'warning', 'Warning');
      return; 
    }
   
    this.activePopover = -1
    this.titleSetting()
    this.setHeaderData()
    this.GetAppointment()
  }

  exportToExcel(): void {
    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('Service Appointments');

    const titleRow = worksheet.addRow(['Service Appointments']);
    titleRow.eachCell((cell, number) => {
      cell.alignment = {
        indent: 1,
        vertical: 'middle',
        horizontal: 'left',
      };
    });
    titleRow.font = { name: 'Arial', family: 4, size: 12, bold: true };
    titleRow.worksheet.mergeCells('A2', 'D2');

    const PresentYear = this.shared.datePipe.transform(this.FromDate, 'yyyy');
    const FromDate = this.shared.datePipe.transform(this.FromDate, 'dd');
    const ToDate = this.shared.datePipe.transform(this.ToDate, 'dd');
    const PresentMonth = this.shared.datePipe.transform(this.FromDate, 'MMMM');
  
     let storeValue = '';

      if (!this.StoreValues || this.StoreValues.length === 0) {
        // No selection → All stores
        storeValue = this.stores.map((s: any) => s.storename).join(', ');
      }
      else if (this.StoreValues.length === this.stores.length) {
        // All selected → bind all store names
        storeValue = this.stores.map((s: any) => s.storename).join(', ');
      }
      else {
        // Partial selection
        storeValue = this.stores
          .filter((s: any) => this.StoreValues.includes(s.ID))
          .map((s: any) => s.storename)
          .join(', ');
      }
      
          const filters = [
            { name: 'Store:', values: storeValue },
    { name: 'Time Frame :', values: this.FromDate + ' to ' + this.ToDate }
    ]
    // const ReportFilter = worksheet.addRow(['Service Appointments']);
    // ReportFilter.font = { name: 'Arial', family: 4, size: 10, bold: true };
    let startIndex = 2
    filters.forEach((val: any) => {
      startIndex++
      worksheet.addRow('');
      worksheet.getCell(`A${startIndex}`);
      worksheet.mergeCells(`B${startIndex}:C${startIndex}`);
      worksheet.getCell(`A${startIndex}`).value = val.name;
      worksheet.getCell(`B${startIndex}`).value = val.values
    })
    worksheet.addRow('');

    const secondHeader = [
      'Dealer', 'Customer', 'Contact Number', 'Date', 'Time', 'NoShow', 'Assigned User', 'Stock Number', 'Year', 'Make', 'Model',
      'Vin', 'Mileage'
    ];
    worksheet.addRow(secondHeader);
    worksheet.getRow(6).height = 25;
    worksheet.getRow(6).font = { bold: true, color: { argb: 'FFFFFFFF' } };
    worksheet.getRow(6).alignment = { vertical: 'middle', horizontal: 'center' };
    worksheet.getRow(6).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2F5597' } };

    const bindingHeaders = [
      'AS_COMPANYNAME', 'CUSTOMER', 'CONTACT NUMBER', 'APPTDate', 'APPOINTMENT TIME', 'NO SHOW',
      'ASSIGNED USER', 'STOCK NUMBER', 'MODELYEAR', 'MAKE', 'MODEL', 'VIN',
      'MILEAGE'];
    const currencyFields: any = [];
    for (const info of this.Appointmentdata) {
      const rowData = bindingHeaders.map(key => {
        const val = info[key];
        return key == 'NO SHOW' ? (val == 'Y' ? 'Y' : '-') : (val === 0 || val == null || val == '' ? '-' : val);
      });

      const dealerRow = worksheet.addRow(rowData);
      // dealerRow.font = { bold: false };

      bindingHeaders.forEach((key, index) => {
        const cell = dealerRow.getCell(index + 1);
        if (currencyFields.includes(key) && typeof cell.value === 'number') {
          cell.numFmt = '"$"#,##0.00';
          cell.alignment = { horizontal: 'right' };
        } else if (!isNaN(Number(cell.value))) {
          cell.alignment = { horizontal: 'center' };
        }
      });
    }
    worksheet.columns.forEach((column: any) => {
      let maxLength = 15;
      column.width = maxLength + 2;
    });
    workbook.xlsx.writeBuffer().then((buffer: any) => {
      const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      this.shared.exportToExcel(workbook, 'Service Appointments')
    });
  }
}
