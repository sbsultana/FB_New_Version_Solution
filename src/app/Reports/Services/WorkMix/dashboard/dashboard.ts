import { Component, Injector, HostListener, OnInit, AfterViewInit } from '@angular/core';
import { Api } from '../../../../Core/Providers/Api/api';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { NgbActiveModal } from '@ng-bootstrap/ng-bootstrap';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
import { BsDatepickerConfig, BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { Stores } from '../../../../CommonFilters/stores/stores';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { Setdates } from '../../../../Core/Providers/SetDates/setdates';
import { DateRangePicker } from '../../../../CommonFilters/date-range-picker/date-range-picker';
import { NgxSpinnerService } from 'ngx-spinner';
import * as FileSaver from 'file-saver';
import { DatePipe } from '@angular/common';
import { Subscription } from 'rxjs';
import { Workbook } from 'exceljs';
const EXCEL_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
const EXCEL_EXTENSION = '.xlsx';

@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [SharedModule, BsDatepickerModule, Stores, DateRangePicker],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
})
export class Dashboard {
  // Properties
  LMstate: any;
  excel!: Subscription;
  FromDate: any = '';
  ToDate: any = '';
  tempDateType: string = '';
  tempDisplayTime: string = '';
  bsRangeValue: any[] = [];
  custom: boolean = false;
  DateType: any = 'MTD';
  displaytime: any = '';
  toporbottom: 'T' | 'B' = 'B';
  WorkMix: any[] = [];
  WorkMixTotalsData: any[] = [];
  NoData: any = ''
  expanddataarray: any = [];
  stores: any = []
  groupsArray: any = [];
  storename: any = ''
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  groupId: any = 0;
  storeIds: any = 0;
  activePopover = -1;
  month!: Date;
  minDate!: Date;
  maxDate!: Date;
  isDesc = false;
  column = '';
  zeroHours: string = 'Y';
  reporttotals: any = 'B'
  storesFilterData: any = {
    'groupsArray': this.groupsArray, 'groupId': this.groupId, 'storesArray': this.stores, 'storeids': '1', 'type': 'M', 'others': 'N',
    'groupName': this.groupName, 'storename': this.storename, storecount: null, 'storedisplayname': this.storedisplayname
  };
  Dates: any = {
    'FromDate': this.FromDate, 'ToDate': this.ToDate, "MaxDate": this.maxDate, 'MinDate': this.minDate, 'DateType': this.DateType, 'DisplayTime': this.displaytime,
    Types: [
      { 'code': 'MTD', 'name': 'MTD' },
      // { 'code': 'QTD', 'name': 'QTD' },
      { 'code': 'YTD', 'name': 'YTD' },
      { 'code': 'PYTD', 'name': 'PYTD' },
      // { 'code': 'LY', 'name': 'Last Year' },
      { 'code': 'LM', 'name': 'Last Month' },
      { 'code': 'PM', 'name': 'Same Month PY' },
    ]
  }
  DupMonth: Date = new Date();
  DuplicatDate!: Date;
  bsConfig: Partial<BsDatepickerConfig> = {
    dateInputFormat: 'MMMM/YYYY',
    minMode: 'month'
  };
  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent) {
    const clickedInside = (event.target as HTMLElement).closest('.dropdown-toggle, .reportstores-card , .timeframe');
    if (!clickedInside) {
      this.activePopover = -1;
    }
  }
  // Constructor

  constructor(public apiSrvc: Api, public shared: Sharedservice, public setdates: Setdates, private toast: ToastService,
    public spinner: NgxSpinnerService, private datepipe: DatePipe,) {

    let today = new Date();

    // current month (no subtraction)
    this.month = new Date(today.getFullYear(), today.getMonth(), 1);

    // maxDate → current month
    this.maxDate = new Date(today.getFullYear(), today.getMonth(), 1);

    // minDate → 3 years back
    this.minDate = new Date();
    this.minDate.setFullYear(today.getFullYear() - 3);
    this.maxDate = new Date();
    this.minDate = new Date();
    this.minDate.setFullYear(this.maxDate.getFullYear() - 3);
    this.maxDate.setMonth(this.maxDate.getMonth() - 1);
    if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
      // this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.ustores.split(',')
      this.groupId = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Preferences
      this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids.split(',')
    }
    console.log('store displayname', this.storedisplayname);
    if (this.shared.common.groupsandstores.length > 0) {
      this.groupsArray = this.shared.common.groupsandstores.filter((val: any) => val.sg_id != this.shared.common.reconID);
      this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
      this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_Name : this.groupName = ''
      this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
      this.getStoresandGroupsValues()
    }
    this.shared.setTitle(this.shared.common.titleName + '-WorkMix')

    this.setHeaderData()
    this.initializeDates('MTD'); // default
    this.setDates(this.DateType); // UI display
    this.GetWorkMix();
  }

  //Reports Totals & Zero hours
  multipleorsingle(block: any, e: any) {

    // ZERO HOURS TOGGLE
    if (block === 'ZH') {
      this.zeroHours = e; // 'Y' or 'N'
    }

    //  TOP/BOTTOM FIX (you also had bug here)
    if (block === 'TB') {
      this.toporbottom = e; // 'T' or 'B'
    }
    if (block === 'ZH') {
      this.zeroHours = e;
      this.activePopover = -1; //  close dropdown
    }
  }
  ExcelStoreNames: any = [];
  exportToExcel() {

    // HANDLE storeIds (array OR string)
    const storeIdsArr = Array.isArray(this.storeIds)
      ? this.storeIds.map((x: any) => x.toString())
      : this.storeIds?.toString().split(',');

    // GET GROUP DATA
    const groupData = this.shared.common.groupsandstores
      .find((v: any) => v.sg_id == this.groupId);

    if (!groupData) {
      this.toast.show('Group not found', 'warning', 'Warning');
      return;
    }

    // FILTER SELECTED STORES
    const selectedStores = groupData.Stores.filter((item: any) =>
      storeIdsArr.includes(item.ID.toString())
    );

    // SHOW "ALL STORES" OR NAMES
    if (selectedStores.length === groupData.Stores.length) {
      this.ExcelStoreNames = 'All Stores';
    } else {
      this.ExcelStoreNames = selectedStores.map((a: any) => a.storename);
    }

    const data = this.WorkMix || [];

    if (!data.length) {
      this.toast.show('No data available', 'warning', 'Warning');
      return;
    }

    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet('Work Mix');

    worksheet.views = [{ showGridLines: false }];

    worksheet.addRow('');

    //  TITLE
    const titleRow = worksheet.addRow(['Work Mix']);
    titleRow.font = { name: 'Arial', size: 12, bold: true };
    titleRow.alignment = { horizontal: 'left', indent: 1 };
    worksheet.mergeCells('A2', 'D2');

    worksheet.addRow('');

    const DateToday = this.datepipe.transform(new Date(), 'MM/dd/yyyy h:mm:ss a');
    const DATE_EXTENSION = this.datepipe.transform(new Date(), 'MMddyyyy');

    worksheet.addRow([DateToday]).font = { size: 9 };

    // FILTER SECTION
    worksheet.addRow(['Report Controls :']).font = { bold: true };

    worksheet.getCell('B8').value = 'Group :';
    worksheet.getCell('D8').value = groupData.sg_name;

    worksheet.getCell('B9').value = 'Stores :';

    worksheet.mergeCells('D9', 'O12');

    const storesCell = worksheet.getCell('D9');

    storesCell.value = Array.isArray(this.ExcelStoreNames)
      ? this.ExcelStoreNames.join(', ')   // partial stores
      : this.ExcelStoreNames;             // "All Stores"

    storesCell.alignment = {
      wrapText: true,
      vertical: 'top',
      horizontal: 'left'
    };

    worksheet.addRow('');

    // HEADERS (MATCH UNAPPLIED STYLE)
    const headers = [
      'Store Name',
      'Hours',
      'Hours %',
      'Sales',
      'Sales %',
      'Avg Sales',
      'ELR'
    ];

    const headerRow = worksheet.addRow(headers);

    headerRow.eachCell((cell: any) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '2a91f0' }
      };

      cell.font = {
        color: { argb: 'FFFFFF' },
        bold: true
      };

      cell.alignment = {
        vertical: 'middle',
        horizontal: 'center'
      };

      //  DOTTED BORDER
      cell.border = {
        right: { style: 'dotted' }
      };
    });

    //  DATA
    for (const d of data) {

      // STORE ROW
      const storeRow = worksheet.addRow([
        d.Store_Name || '-', '', '', '', '', '', ''
      ]);

      storeRow.font = { size: 9 };

      storeRow.getCell(1).alignment = {
        horizontal: 'left',
        indent: 1
      };

      // ZEBRA ROW
      if (storeRow.number % 2) {
        storeRow.eachCell((cell: any) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'e5e5e5' }
          };
        });
      }

      // SUB ROWS
      if (d.sub) {
        for (const s of d.sub) {

          const subRow = worksheet.addRow([
            s.Paytype || '-',
            s.Hours ?? '-',
            s.HoursPer ? s.HoursPer + '%' : '-',
            s.Sale ?? '-',
            s.SalePer ? s.SalePer + '%' : '-',
            s.SaleAvg ?? '-',
            s.ELR ?? '-'
          ]);

          subRow.font = { size: 9 };

          subRow.getCell(1).alignment = {
            horizontal: 'left',
            indent: 2
          };

          subRow.eachCell((cell: any, col: number) => {

            //  DOTTED BORDER
            cell.border = {
              top: { style: 'thin' },
              left: { style: 'thin' },
              bottom: { style: 'thin' },
              right: { style: 'thin' }
            };

            //  FORMATTING
            if (col === 4 || col === 6 || col === 7) {
              cell.numFmt = '$#,##0';
            }

            if (col === 2) {
              cell.numFmt = '#,##0.00';
            }

            if (col !== 1) {
              cell.alignment = {
                horizontal: 'right'
              };
            }
          });

          // ZEBRA
          if (subRow.number % 2) {
            subRow.eachCell((cell: any) => {
              cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'e5e5e5' }
              };
            });
          }
        }
      }
    }

    // COLUMN WIDTHS
    worksheet.getColumn(1).width = 30;

    for (let i = 2; i <= 10; i++) {
      worksheet.getColumn(i).width = 12;
    }

    worksheet.addRow('');

    // DOWNLOAD
    workbook.xlsx.writeBuffer().then((data: any) => {

      const blob = new Blob([data], {
        type: EXCEL_TYPE
      });

      FileSaver.saveAs(
        blob,
        'Work Mix_' + DATE_EXTENSION + EXCEL_EXTENSION
      );
    });
  }

  //getstores
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

  //apply filters
  viewreport() {
    this.activePopover = -1;

    if (!this.storeIds || this.storeIds.length === 0) {
      this.toast.show('Please Select Atleast One Store', 'warning', 'Warning');
      return;
    }

    // ✅ APPLY TEMP VALUES HERE
    if (this.tempDateType) {
      this.FromDate = this.FromDate;
      this.ToDate = this.FromDate;
      this.DateType = this.tempDateType;
      this.displaytime = this.tempDisplayTime;
    }

    const data = {
      Reference: 'Work Mix',
      FromDate: this.FromDate,
      ToDate: this.ToDate,
      storeValues: this.storeIds.toString(),
      dateType: this.DateType,
      groups: this.groupId,
    };

    this.shared.api.SetReports({ obj: data });

    this.GetWorkMix(); // ✅ API call only here
  }

  ngAfterViewInit() {

    this.apiSrvc.getStores().subscribe((res: any) => {
      if (this.shared.common.pageName == 'Work Mix') {
        if (res.obj.storesData != undefined) {
          this.groupsArray = res.obj.storesData;
          this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
          this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_name : this.groupName = ''
          this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
          this.getStoresandGroupsValues()
        }
      }
    });
    this.excel = this.apiSrvc.getExportToExcelAllReports().subscribe((res) => {
      if (this.excel != undefined) {
        this.LMstate = res.obj.state;
        if (res.obj.title == 'Work Mix') {
          if (res.obj.state == true) {
            this.exportToExcel();
          }
        }
      }
    });
  }

  //Time frame dates

  setDates(type: any) {
    this.displaytime = '(' + this.Dates.Types.find((val: any) => val.code == type)?.name + ')';

    this.maxDate = new Date();
    this.minDate = new Date();
    this.minDate.setFullYear(this.maxDate.getFullYear() - 3);

    this.Dates.FromDate = this.FromDate;
    this.Dates.ToDate = this.ToDate;
    this.Dates.DateType = this.DateType;
    this.Dates.DisplayTime = this.displaytime;
  }
  onTimeframeChange(type: any) {
    this.initializeDates(type);
    this.setDates(type); // UI
  }
  initializeDates(type: any) {
    const dates = this.setdates.setDates(type);

    this.FromDate = dates[0] ?? '';
    this.ToDate = dates[1] ?? '';

    this.DateType = type;
  }
  updatedDates(data: any) {
    this.FromDate = data.FromDate;
    this.ToDate = data.ToDate;
    this.DateType = data.DateType;
    this.displaytime = data.DisplayTime;
  }
  dateRangeCreated($event: any) {
    if ($event) {
      this.FromDate = this.shared.datePipe.transform($event[0], 'MM-dd-yyyy');
      this.ToDate = this.shared.datePipe.transform($event[1], 'MM-dd-yyyy');

      this.DateType = 'C'; // optional (only for UI)
      this.custom = true;
    }
  }
  onOpenCalendar(container: any) {
    container.setViewMode('month');
    container.monthSelectHandler = (event: any): void => {
      container.value = event.date;
      this.month = event.date;
    };
  }

  changeDate(date: Date) {
    this.month = date;
  }

  // Header
  setHeaderData() {
    this.apiSrvc.SetHeaderData({
      obj: {
        title: 'Work Mix',
        stores: this.storeIds,
        groups: this.groupId,
        month: this.month
      }
    });
  }

  // API Calls
  GetWorkMix() {
    this.WorkMix = [];
    this.NoData = false;
    this.shared.spinner.show();

    const payload = {
      StartDate: this.FromDate,
      EndDate: this.ToDate,
      Stores: this.storeIds.join(','),
      ExcludeZeroHours: this.zeroHours,
    };

    console.log('Payload:', payload);

    this.apiSrvc
      .postmethod(this.shared.common.routeEndpoint + 'GetServiceWorkMix', payload)
      .subscribe({
        next: (res: any) => {
          this.shared.spinner.hide();

          if (res.status !== 200 || !res.response) {
            this.NoData = true;
            return;
          }

          const grouped = this.groupByStore(res.response);

          const reportTotal = grouped.filter(
            (e: any) => e.Store_Name.toLowerCase() === 'report total'
          );

          const data = grouped.filter(
            (e: any) => e.Store_Name.toLowerCase() !== 'report total'
          );

          this.WorkMix =
            this.toporbottom === 'T'
              ? [...reportTotal, ...data]
              : [...data, ...reportTotal];

          this.NoData = this.WorkMix.length === 0;
        },
        error: () => {
          this.shared.spinner.hide();
          this.NoData = true;
        },
      });
  }


  private groupByStore(data: any[]) {
    return data.reduce((result: any[], item: any) => {
      const exists = result.find(r => r.Store_Name === item.Store_Name);

      if (!exists) {
        result.push({
          Store_Name: item.Store_Name,
          sub: data.filter(v => v.Store_Name === item.Store_Name)
        });
      }

      return result;
    }, []);
  }

  // UI Helpers

  inTheGreen(value: number): boolean {
    return value >= 0;
  }

  togglePopover(index: number) {
    this.activePopover = this.activePopover === index ? -1 : index;
  }
}