import { Component, OnInit, ViewChild, ElementRef, HostListener } from '@angular/core';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { Setdates } from '../../../../Core/Providers/SetDates/setdates';
import { BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { common } from '../../../../common';
// import { Stores } from '../../../../CommonFilters/stores/stores';
import { DateRangePicker } from '../../../../CommonFilters/date-range-picker/date-range-picker';
import { Subscription } from 'rxjs';
import { Stores } from '../../../../CommonFilters/stores/stores';
import { FilterPipe } from '../../../../Core/Providers/filterpipe/filter.pipe';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [SharedModule, BsDatepickerModule, FilterPipe, DateRangePicker, Stores],
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

  // groups: any = 1
  filteredCustomers: any;
  Appointmentsearch: any;
  searcdata: any;
  FromDate: any = '';
  ToDate: any = '';

  minDate!: Date;
  maxDate!: Date;
  todaytitle: any;
  reportOpenSub!: Subscription;
  reportGetting!: Subscription;
  Pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;

  // FromDate: any;
  // ToDate: any;
  DateType: string = 'C';
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
  StoreValues: any = '0';
  popup: any = [{ type: 'Popup' }];
  groups: any = 1;
  gridvisibility: any;
  bsRangeValue!: Date[];
  groupsArray: any = [];
  storename: any = ''
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  groupId: any = 0;
  stores: any = []
  // activePopover: number = -1;
  storesFilterData: any = {
    'groupsArray': this.groupsArray, 'groupId': this.groupId, 'storesArray': this.stores, 'storeids': '1', 'type': 'M', 'others': 'N',
    'groupName': this.groupName, 'storename': this.storename, storecount: null, 'storedisplayname': this.storedisplayname
  };
  // @HostListener('document:click', ['$event'])
  // onDocumentClick(event: MouseEvent) {
  //   const clickedInside = (event.target as HTMLElement).closest('.dropdown-toggle, .dropdown-menu , .timeframe, .reportstores-card');
  //   if (!clickedInside) {
  //     this.activePopover = -1;
  //   }
  // }
  constructor(public shared: Sharedservice, public setdates: Setdates, private toast: ToastService,

  ) {
    if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
      // this.StoreValues = '1,2';
      this.groupId = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Preferences
      this.StoreValues = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids.split(',')
      if (this.StoreValues.toString().indexOf(',') > 0) {
        this.gridvisibility = 'DL';
      } else {
        this.gridvisibility = 'SL';
      }
    }

    localStorage.setItem('time', 'C');

    this.titleSetting()
    this.shared.setTitle(this.shared.common.titleName + '- Sales Appointments');

    this.todaytitle = this.shared.datePipe.transform(new Date(), 'MM/dd/yyyy');
    let today = new Date();
    const tomorrow = new Date();
    tomorrow.setDate(today.getDate());
    this.FromDate = this.shared.datePipe.transform(today, 'MM/dd/yyyy');
    this.ToDate = this.shared.datePipe.transform(tomorrow, 'MM/dd/yyyy');
    this.setDates(this.DateType)
    // if (localStorage.getItem('Fav') != 'Y') {
    const data = {
      title: this.dynamicTitle,
      stores: this.StoreValues,
      groups: this.groups,
      count: 0,
      fromdate: this.FromDate,
      todate: this.ToDate,
    };
    this.shared.api.SetHeaderData({
      obj: data,
    });

    this.GetAppointment();
    // }
  }
  dynamicTitle = 'Sales Appointments'
  titleSetting() {
    // if (localStorage.getItem('time') == 'C') {
    console.log(this.todaytitle, this.shared.datePipe.transform(this.FromDate, 'MM/dd/yyyy'), this.shared.datePipe.transform(this.ToDate, 'MM/dd/yyyy'));

    // if ((this.todaytitle == this.shared.datePipe.transform(this.FromDate, 'MM/dd/yyyy')) && (this.todaytitle == this.shared.datePipe.transform(this.ToDate, 'MM/dd/yyyy'))) {
    //   this.dynamicTitle = 'Today Sales Appointments'
    // }
    // else {
    this.dynamicTitle = 'Sales Appointments'
    // }
    const data = {
      title: this.dynamicTitle,
      stores: this.StoreValues,
      groups: this.groups,
      count: 0,
      fromdate: this.FromDate,
      todate: this.ToDate,
    };
    this.shared.api.SetHeaderData({
      obj: data,
    });

    // }
  }

  setHeaderData() {
    const data = {
      title: this.dynamicTitle,
      stores: this.storeIds,
      groups: 1,
      count: 0,
      fromdate: this.FromDate,
      todate: this.ToDate,
    };
    this.shared.api.SetHeaderData({
      obj: data,
    });
  }
  ngOnInit(): void {
    this.getStores();
  }


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
  getStoresandGroupsValues() {
    this.storesFilterData.groupsArray = this.groupsArray;
    this.storesFilterData.groupId = this.groupId;
    this.storesFilterData.storesArray = this.stores;
    this.storesFilterData.storeids = this.StoreValues;
    this.storesFilterData.groupName = this.groupName;
    this.storesFilterData.storename = this.storename;
    this.storesFilterData.storecount = this.storecount;
    this.storesFilterData.storedisplayname = this.storedisplayname;

    this.storesFilterData = {
      groupsArray: this.groupsArray,
      groupId: this.groupId,
      storesArray: this.stores,
      storeids: this.StoreValues,
      groupName: this.groupName,
      storename: this.storename,
      storecount: this.storecount,
      storedisplayname: this.storedisplayname,
      'type': 'M', 'others': 'N'
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
    this.storedisplayname = data.storedisplayname;
  }
  getStores() {
    this.selectedstorevalues = [];
  }

  GetAppointment() {
    this.Appointmentdata = [];
    this.NoData = false;
    this.shared.spinner.show();
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
      this.storeIds = InvObj.dataobj.Data1
      console.log(this.todaytitle);

      // if ((this.todaytitle == this.shared.datePipe.transform(this.FromDate, 'MM/dd/yyyy')) && (this.todaytitle == this.shared.datePipe.transform(this.ToDate, 'MM/dd/yyyy'))) {
      //   this.dynamicTitle = 'Today Sales Appointments'
      // }
      // else {
      this.dynamicTitle = 'Sales Appointments'
      // }
      const data = {
        title: this.dynamicTitle,

        stores: this.StoreValues,
        groups: 1,
        count: 0,
        fromdate: this.FromDate,
        todate: this.ToDate,
      };
      this.shared.api.SetHeaderData({
        obj: data,
      });

    } else {
      obj = {

        DealerId: this.StoreValues,
        StartDate: this.shared.datePipe.transform(this.FromDate, 'MM/dd/yyyy'),
        EndDate: this.shared.datePipe.transform(this.ToDate, 'MM/dd/yyyy'),
      };
      console.log(obj);

    }
    // const curl = this.shared.getEnviUrl() +this.shared.common.routeEndpoint+'GetDrivecentricAppointments';
    this.shared.api.postmethod(this.shared.common.routeEndpoint + 'GetVSAppointmentsData', obj).subscribe(
      (res: { message: any; status: number; response: string | any[] | undefined; }) => {
        const currentTitle = document.title;
        // this.shared.api.logSaving(curl, {}, '', res.message, currentTitle);
        if (res.status == 200) {
          this.Appointmentdata = []
          this.Appointmentsearch = []
          this.appointment = []
          if (res.response != undefined) {
            if (res.response.length > 0) {
              this.shared.spinner.hide();
              this.Appointmentdata = [...res.response];
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

          localStorage.removeItem('DetailsObject');
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
      if (this.shared.common.pageName == 'Sales Appointments') {
        if (res.obj.storesData != undefined) {
          this.groupsArray = res.obj.storesData;
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
            // this.sendEmailData(res.obj.Email, res.obj.notes, res.obj.from);
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
    this.DateType == 'C' ? this.displaytime = ' (  ' + this.shared.datePipe.transform(this.FromDate, 'MM/dd/yyyy') + ' - ' + this.shared.datePipe.transform(this.ToDate, 'MM/dd/yyyy') + ' ) ' :
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

      this.toast.show('Please select at least one store', 'warning', 'Warning');
      return;
    }

    this.activePopover = -1
    this.titleSetting()
    this.setHeaderData()
    this.GetAppointment()

  }
  getExcelStoreLabel(): string {
    // 0 selected
    if (!this.StoreValues || this.StoreValues.length === 0) {
      return 'Selected (0)';
    }

    // All selected → show all store names
    if (this.StoreValues.length === this.stores.length) {
      return this.stores
        .map((s: any) => s.storename)
        .join(', ');
    }

    // Single selected → store name
    if (this.StoreValues.length === 1) {
      const store = this.stores.find(
        (s: any) => s.ID === this.StoreValues[0]
      );
      return store ? store.storename : '-';
    }

    // Multiple (not all)
    return `Selected (${this.StoreValues.length})`;
  }


  ExcelStoreNames: any = []
  exportToExcel(): void {
    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('Sales Appointments');

    // const groupObj = this.shared.common.groupsandstores.find((v: any) => v.sg_id == this.groups);
    // if (!groupObj) return;


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
    console.log(this.ExcelStoreNames, 'exclstoreNames');
    const filters = [
      { name: 'Stores:', values: storeValue },
      { name: 'Time Frame :', values: this.FromDate + ' to ' + this.ToDate }
    ];

    let currentRow = worksheet.lastRow?.number ?? worksheet.rowCount;
    filters.forEach((filter) => {
      currentRow++;

      let value = Array.isArray(filter.values)
        ? filter.values.join(', ')
        : filter.values;

      const row = worksheet.addRow([filter.name, value]);
      row.getCell(1).font = { bold: true, name: 'Arial', size: 10 };
      row.getCell(2).font = { name: 'Arial', size: 10, color: { argb: 'FF1F497D' } }; // blue color for values
      worksheet.mergeCells(`B${currentRow}:F${currentRow}`);
    });


    worksheet.addRow('');

    const secondHeader = [
      'Dealer', 'Time', 'Date', 'Stock Number', 'Year', 'Make', 'Model', 'Customer',
      'Assigned User', 'Appt ID', 'Appt Status', 'Lead Type', 'Status',
      'Sales Representative', 'BD Agent', 'Manager', 'Created', 'Source', 'Contacted'
    ];

    const headerRow = worksheet.addRow(secondHeader);
    const headerRowNumber = headerRow.number;


    headerRow.eachCell((cell: any) => {
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF2F5597' }
      };
    });


    headerRow.height = 25;

    const bindingHeaders = [
      'Dealer', 'Appt Start Time', 'Appt Start Date', 'Stock Number', 'Year', 'Make',
      'Model', 'Customer', 'Assigned User', 'Appt ID', 'Appt Status', 'Lead Type',
      'Status', 'Sales Representative', 'BD Agent', 'Manager', 'Created', 'Source', 'Contacted'
    ];

    for (const info of this.Appointmentdata) {

      const rowData = bindingHeaders.map(key => {
        let val = info[key];


        if (key === 'Appt Start Date' || key === 'Created') {
          if (val) {
            const d = new Date(val);
            const mm = String(d.getMonth() + 1).padStart(2, '0');
            const dd = String(d.getDate()).padStart(2, '0');
            const yyyy = d.getFullYear();
            val = `${mm}/${dd}/${yyyy}`;
          } else {
            val = '-';
          }
        }

        return (val === 0 || val == null || val === '') ? '-' : val;
      });

      const dealerRow = worksheet.addRow(rowData);

      dealerRow.eachCell((cell: any) => {
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
      });
    }

    worksheet.columns.forEach((column: any) => {
      column.width = 18;
    });

    workbook.xlsx.writeBuffer().then((buffer: any) => {
      const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      this.shared.exportToExcel(workbook, 'Sales Appointments');
    });
  }

}




