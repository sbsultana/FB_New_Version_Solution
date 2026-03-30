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
@Component({
  selector: 'app-dashboard',
  imports: [SharedModule, BsDatepickerModule, DateRangePicker, Stores],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
})
export class Dashboard {
  ServiceData: any = [];
  IndividualServiceGross: any = [];
  TotalServiceGross: any = [];

  NoData: boolean = false;
  Department: any = ['Parts'];
  Paytype: any = ['Customerpay', 'Warranty', 'Internal'];




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
  otherStoresArray: any = [];
  otherStoreIds: any = [];

  storesFilterData: any = {
    'groupsArray': this.groupsArray, 'groupId': this.groupId, 'storesArray': this.stores, 'storeids': '1', 'type': 'M', 'others': 'Y',
    'groupName': this.groupName, 'storename': this.storename, storecount: null, 'storedisplayname': this.storedisplayname,
    otherStoresArray: this.otherStoresArray, otherStoreIds: this.otherStoreIds

  };

  FromDate: any = '';
  ToDate: any = '';
  minDate!: Date;
  maxDate!: Date;
  DateType: any = 'MTD';
  displaytime: any = '';
DupFromDate: any = '';
  DupToDate: any = ''

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
    if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
      this.groupId = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Preferences
      this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids.split(',')
      this.otherStoreIds = JSON.parse(localStorage.getItem('otherstoreids')!);

    }
    if (this.shared.common.groupsandstores.length > 0) {
      this.groupsArray = this.shared.common.groupsandstores.filter((val: any) => val.sg_id != this.shared.common.reconID);
      this.otherStoresArray = this.shared.common.OtherStoresData[0].Stores

      this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
      this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_Name : this.groupName = ''
      this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
      this.getStoresandGroupsValues()
    }

    this.shared.setTitle(this.comm.titleName + '-Parts Summary GL');
    let today = new Date();
    if (today.getDate() == 1) {
      this.initializeDates('LMGL')
    } else if (today.getDate() > 1 && today.getDate() < 5) {
      this.initializeDates('LMGL')
    } else {
      this.initializeDates('MTD')
    }

    this.setHeaderData();
    this.getSubTypeDetail('FL')
    this.GetData();

  }
  asoftime: any = [];
  ngOnInit() { }

  datetype() {
    if (this.DateType == 'PM') {
      return 'SP';
    }
    else if (this.DateType == 'C') {
      return 'C'
    }
    return this.DateType;
  }
  lastyearDate: any;
  setHeaderData() {
    const data = {
      title: 'Parts Summary GL',
      stores: this.storeIds,
      Department: this.Department,
      Paytype: this.Paytype,
      fromdate: this.FromDate,
      todate: this.ToDate,
      groups: this.groupId,
      subtype: this.selectedSubType,
      otherstoreids: this.otherStoreIds

    };
    this.shared.api.SetHeaderData({
      obj: data,
    });

  }
  initializeDates(type: any) {
    let dates: any = this.setdates.setDates(type)
    this.FromDate = dates[0];
    this.ToDate = dates[1];
    type == 'LMGL' ? localStorage.setItem('time', 'LM') : localStorage.setItem('time', type);
    this.lastyearDate = new Date(this.FromDate)
    this.lastyearDate.setFullYear(this.lastyearDate.getFullYear() - 1);
    this.setDates(this.DateType)
  }
  GetData() {
    this.DupFromDate = this.FromDate;
    this.DupToDate = this.ToDate
    this.shared.spinner.show()
    this.IndividualServiceGross = [];
    const obj = {
      "Startdate": this.FromDate.replaceAll('/', '-'),
      "Enddate": this.ToDate.replaceAll('/', '-'),
      "AS_IDS": [...this.storeIds, ...this.otherStoreIds],
      "DEPTTYPE": 'Parts',
      "SUBTYPE": this.selectedSubType.toString(),
      "Details": this.Department.indexOf('Details') >= 0 ? 'Details' : ''

    };
    const curl = environment.apiUrl + this.comm.routeEndpoint + 'GetServiceGrossGLSummaryReport';
    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetServiceGrossGLSummaryReport', obj).subscribe(
      (res) => {
        const currentTitle = document.title;
        this.shared.api.logSaving(curl, {}, '', res.message, currentTitle);
        if (res.status == 200) {
          if (res.response != undefined) {
            if (res.response.length > 0) {
              console.log(res.response);
              this.IndividualServiceGross = res.response;
              this.NoData = false;
              this.IndividualServiceGross.some(function (x: any) {
                if (x.data2 != undefined && x.data2 != null) {
                  if (x.data2.length > 0) {
                    x.Data2 = JSON.parse(x.data2);
                    x.Data2 = x.Data2.map((v: any) => ({ ...v, data2sign: '-', }));
                  }
                }
              });

              this.ServiceData = this.IndividualServiceGross.reduce((r: any, { STORE, ...rest }: any) => {
                if (!r.some((o: any) => o.STORE == STORE)) {
                  r.push({
                    STORE,
                    ...rest,
                    Dealer: '-',
                    Subdata: this.IndividualServiceGross.filter((v: any) => v.STORE == STORE),
                  });
                }
                return r;
              }, []);

              console.log(this.ServiceData);
              this.shared.spinner.hide();
              this.NoData = false;


            } else {

              this.shared.spinner.hide();
              this.NoData = true;
            }
          } else {

            this.shared.spinner.hide();
            this.NoData = true;
          }
        } else {
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
  sort(property: any, data: any) {
    this.isDesc = !this.isDesc; //change the direction
    this.column = property;
    let direction = this.isDesc ? 1 : -1;
    // //console.log(property)
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

  expandorcollapse(ind: any, e: any, ref: any, Item: any, parentData: any) {
    let id = (e.target as Element).id;
    if (id == 'D_' + ind) {
      if (ref == '-') {
        Item.Dealer = '+';
      }
      if (ref == '+') {
        Item.Dealer = '-';
      }
    }
    console.log(ind, id, Item.data2sign);
    if (id == 'DVN_' + ind) {
      console.log(ind, id, Item.data2sign, 'Internal');

      if (ref == '-') {
        Item.data2sign = '+';
      }
      if (ref == '+') {
        Item.data2sign = '-';
      }
    }
  }

  ngAfterViewInit() {
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.comm.pageName == 'Parts Summary GL') {
        if (res.obj.storesData != undefined) {
          this.groupsArray = res.obj.storesData;
          this.otherStoresArray = this.shared.common.OtherStoresData[0].Stores

          this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
          this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_name : this.groupName = ''
          this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
          this.getStoresandGroupsValues()
        }
      }
    })

    this.excel = this.shared.api.getExportToExcelAllReports().subscribe((res) => {
      if (this.excel != undefined) {
        if (res.obj.title == 'Parts Summary GL') {
          if (res.obj.state == true) {
            this.exportToExcel();
          }
        }
      }
    });

    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res) => {
      if (this.email != undefined) {
        if (res.obj.title == 'Parts Summary GL') {
          if (res.obj.stateEmailPdf == true) {
            // this.sendEmailData(res.obj.Email, res.obj.notes, res.obj.from);
          }
        }
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res) => {
      if (this.print != undefined) {
        if (res.obj.title == 'Parts Summary GL') {
          if (res.obj.statePrint == true) {
            //  this.GetPrintData();
          }
        }
      }
    });

    this.Pdf = this.shared.api.getExportToPDFAllReports().subscribe((res) => {
      if (this.Pdf != undefined) {
        if (res.obj.title == 'Parts Summary GL') {
          if (res.obj.statePDF == true) {
            //   this.generatePDF();
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

  StoresData(data: any) {
    this.storeIds = data.storeids;
    this.groupId = data.groupId;
    this.storename = data.storename;
    this.groupName = data.groupName;
    this.storecount = data.storecount;
    this.storedisplayname = data.storedisplayname;
    this.otherStoreIds = data.otherStoreIds;
    this.getSubTypeDetail()
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
    this.storesFilterData.otherStoreIds = this.otherStoreIds;
    this.storesFilterData.otherStoresArray = this.otherStoresArray;

    this.storesFilterData = {
      groupsArray: this.groupsArray,
      groupId: this.groupId,
      storesArray: this.stores,
      storeids: this.storeIds,
      groupName: this.groupName,
      storename: this.storename,
      storecount: this.storecount,
      storedisplayname: this.storedisplayname,
      'type': 'M', 'others': 'Y', otherStoresArray: this.otherStoresArray,
      otherStoreIds: this.otherStoreIds
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

  AllSubType(Data: any) {
    let data = Data.SubData.filter((item: any) => !this.selectedSubType.some((itemToBeRemoved: any) => itemToBeRemoved === item.SubtypeParam))
    console.log(data, Data);
    console.log(this.mainSubType.length);

    if (data.length == 0) {
      let a = Data.SubData.map(function (a: any) {
        return a.SubtypeParam;
      });
      this.selectedSubType = this.selectedSubType.filter((val: any) => !a.includes(val));
      const index = this.mainSubType.findIndex((i: any) => i == Data.Dept);
      console.log(index);

      if (index >= 0) {
        console.log(this.mainSubType);
        this.mainSubType.splice(index, 1);
        console.log(this.mainSubType);
        this.toast.show('Please select any one subtype from ' + Data.Dept + ' Department', 'warning', 'Warning')
        this.alertState = true;
      }
      this.allSubTypeSelection()
    } else {
      this.mainSubType.push(Data.Dept)
      let a = Data.SubData.map(function (a: any) {
        return a.SubtypeParam;
      });
      let subdata = [...this.selectedSubType, ...a]
      let s = new Set(subdata);
      this.selectedSubType = [...s]
      this.alertState = false;
      this.allSubTypeSelection()
    }

    console.log(this.mainSubType, this.alertState);

  }
  multipleorsingle(block: any, e: any, Data?: any) {
    if (block == 'Dept') {
      const index = this.Department.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.Department.splice(index, 1);
        if (this.Department.length == 0) {
          this.toast.show('Please Select Atleast One Department', 'warning', 'Warning');
        }
      } else {
        this.Department.push(e);
      }
      this.activePopover = 5
      this.getSubTypeDetail('DC')
    }
    if (block == 'ST') {
      const index = this.selectedSubType.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.selectedSubType.splice(index, 1);
        let data = Data.SubData.filter((item: any) => !this.selectedSubType.some((itemToBeRemoved: any) => itemToBeRemoved === item.SubtypeParam))
        console.log(data);
        if (data.length > 0) {
          const mainindex = this.mainSubType.findIndex((i: any) => i == Data.Dept);
          this.mainSubType.splice(mainindex, 1);
        }

        if (this.selectedSubType.length == 0 && this.subTypeDetails.length > 0) {
          this.toast.show('Please Select Atleast One Sub Type From Each Department', 'warning', 'Warning');
          this.alertState = true;
        } else if (data.length == Data.SubData.length) {
          this.toast.show('Please select any one subtype from ' + Data.Dept + ' Department')
          this.alertState = true;
        } else {
          this.alertState = false;
        }
        this.allSubTypeSelection()
      } else {
        this.selectedSubType.push(e);
        this.alertState = false;
        let data = Data.SubData.filter((item: any) => !this.selectedSubType.some((itemToBeRemoved: any) => itemToBeRemoved === item.SubtypeParam))
        if (data.length == 0) {
          this.mainSubType.push(Data.Dept)
          this.alertState = false;
        }
        this.allSubTypeSelection()
      }
      console.log(this.mainSubType, this.alertState, this.selectedSubType, '.....................');

    }

  }

  subTypeDetails: any = [];
  selectedSubType: any = [];
  spinnerLoader: boolean = false;
  setSubTypeOverallData: any = [];
  mainSubType: any = []
  singleSubType: any = '';
  subtypeaction: any = '';
  alertState: boolean = false;

  getSubTypeDetail(data?: any) {
    const obj = {
      "AS_IDS": [...this.storeIds, ...this.otherStoreIds],
      "DEPTTYPE": this.Department.toString()
    }
    this.spinnerLoader = true;
    this.subTypeDetails = [];
    this.subtypeaction = 'Y'
    this.shared.api.postmethod(this.comm.routeEndpoint + '/GetSubtypeDetailTypes', obj).subscribe((res: any) => {
      if (res.status == 200) {
        this.spinnerLoader = false;
        this.subtypeaction = 'N'
        this.setSubTypeOverallData = res.response;
        let a = res.response.map(function (a: any) {
          return a.SubtypeParam;
        });
        // this.mainSubType = [this.department]

        let s = new Set(a);
        this.selectedSubType = [...s]
        this.subTypeDetails = res.response.reduce((r: any, { Dept }: any) => {
          if (!r.some((o: any) => o.Dept == Dept)) {
            r.push({
              Dept,
              SubData: res.response.filter((v: any) => v.Dept == Dept),
            });
          }
          return r;
        }, []);

        this.mainSubType = this.subTypeDetails.map(function (a: any) {
          return a.Dept;
        });




        console.log(this.subTypeDetails, this.mainSubType);

        // else{
        //   this.selectedSubType = []
        // }
      }
    })
  }
  allSubTypeSelection() {
    this.mainSubType = []
    this.subTypeDetails.forEach((val: any) => {
      let a = val.SubData.map(function (a: any) {
        return a.SubtypeParam;
      });
      let data = a.filter((item: any) => !this.selectedSubType.some((itemToBeRemoved: any) => itemToBeRemoved === item))
      console.log(data, this.selectedSubType, a);
      if (data.length <= 0) {
        this.mainSubType.push(val.Dept)
      }
    })
    console.log(this.mainSubType);

  }
  activePopover: number = -1;

  togglePopover(popoverIndex: number) {
    this.activePopover = this.activePopover === popoverIndex ? -1 : popoverIndex;
  }

  viewreport() {
    this.activePopover = -1

    if (this.storeIds.length == 0 && this.otherStoreIds.length == 0) {
      this.toast.show('Please Select Atleast One Store', 'warning', 'Warning');
    } else if (this.Department.length == 0) {
      this.toast.show('Please Select Atleast One Department Type', 'warning', 'Warning');
    }

    else {
      this.setHeaderData()
      this.GetData()
    }

  }
  ExcelStoreNames: any = [];

  exportToExcel() {
    let storeNames: any = [];
    let store = this.storeIds
    // const obj = {
    //   id: this.groupId,
    //   userid: JSON.parse(localStorage.getItem('UserDetails')!).userid,
    // };
    // this.shared.api
    //   .postmethodOne(this.comm.routeEndpoint+'GetStoresbyGroupuserid', obj)
    //   .subscribe((res: any) => {
    storeNames = this.comm.groupsandstores
      .filter((v: any) => v.sg_id == this.groupId)[0]
      .Stores.filter((item: any) =>
        store.some((cat: any) => cat === item.ID.toString())
      );
    if (
      store.length ==
      this.comm.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0]
        .Stores.length
    ) {
      this.ExcelStoreNames = 'All Stores';
    } else {
      this.ExcelStoreNames = storeNames.map(function (a: any) {
        return a.storename;
      });
    }
    const ServiceData = this.ServiceData.map((_arrayElement: any) =>
      Object.assign({}, _arrayElement)
    );
    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('Parts Summary GL');
    worksheet.views = [
      {
        showGridLines: false,
      },
    ];
    worksheet.addRow('');
    const titleRow = worksheet.addRow(['Parts Summary GL']);
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
    const DATE_EXTENSION = this.shared.datePipe.transform(new Date(), 'MMddyyyy');
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

    const timeframe = worksheet.getCell('B7');
    timeframe.value = this.FromDate + ' to ' + this.ToDate;
    timeframe.font = { name: 'Arial', family: 4, size: 9 };
    const Stores = worksheet.addRow(['Stores :']);
    Stores.getCell(1).font = { name: 'Arial', family: 4, size: 9, bold: true };
    const Groups = worksheet.getCell('B8');
    Groups.value = 'Groups :';
    const groups = worksheet.getCell('D8');
    groups.value = this.comm.groupsandstores.filter(
      (val: any) => val.sg_id == this.groupId.toString()
    )[0].sg_name;

    groups.font = { name: 'Arial', family: 4, size: 9 };
    // const Brands = worksheet.getCell('B9');
    // Brands.value = 'Brands :';
    // const brands = worksheet.getCell('D9');
    // brands.value = '-';
    // brands.font = { name: 'Arial', family: 4, size: 9 };
    const Stores1 = worksheet.getCell('B10');
    Stores1.value = 'Stores :';
    const stores1 = worksheet.getCell('D10');
    stores1.value =
      this.ExcelStoreNames == 0
        ? 'All Stores'
        : this.ExcelStoreNames == null
          ? '-'
          : this.ExcelStoreNames.toString().replaceAll(',', ', ');
    stores1.font = { name: 'Arial', family: 4, size: 9 };

    worksheet.addRow('');
    let dateYear = worksheet.getCell('A11');
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
    let count = worksheet.getCell('B11');
    count.value = '';
    count.alignment = { vertical: 'middle', horizontal: 'center' };
    count.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
      color: { argb: 'FFFFFF' },
    };
    count.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '2a91f0' },
      bgColor: { argb: 'FF0000FF' },
    };
    count.border = { right: { style: 'dotted' } };
    worksheet.mergeCells('C11', 'I11');
    let units = worksheet.getCell('C11');
    units.value = 'Sales';
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
    units.border = { right: { style: 'dotted' } };
    worksheet.mergeCells('J11', 'Q11');
    let frontgross = worksheet.getCell('J11');
    frontgross.value = 'Gross';
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

    let Headings = [
      FromDate + ' - ' + ToDate + ', ' + PresentYear,
      'GL CNT',
      'Actual', 'Pace', 'Forecast', 'Var', '% of Forecast', 'Last Year', 'Var', this.datetype() == 'C' ? this.shared.datePipe.transform(this.FromDate, 'MM.dd.yyyy') + '-' + this.shared.datePipe.transform(this.ToDate, 'MM.dd.yyyy') : this.datetype(), 'Pace', 'Forecast', 'Var', '% of Forecast', 'Last Year', 'Var', 'GP %'
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
      var obj = [
        d.STORE == '' ? '-' : d.STORE == null ? '-' : d.STORE,
        d.UNITS == '' ? '-' : d.UNITS == null ? '-' : d.UNITS,
        d.SALES_MTD == '' ? '-' : d.SALES_MTD == null ? '-' : d.SALES_MTD,
        d.SALES_PACE == '' ? '-' : d.SALES_PACE == null ? '-' : d.SALES_PACE,
        d.SALES_TARGET == '' ? '-' : d.SALES_TARGET == null ? '-' : d.SALES_TARGET,
        d.SALES_DIFF == '' ? '-' : d.SALES_DIFF == null ? '-' : d.SALES_DIFF,
        d.SALES_FORECAST == '' ? '-' : d.SALES_FORECAST == null ? '-' : d.SALES_FORECAST + '%',
        d.SALES_LY == '' ? '-' : d.SALES_LY == null ? '-' : d.SALES_LY,
        d.LY_GROSS_DIFF == '' ? '-' : d.LY_GROSS_DIFF == null ? '-' : d.LY_GROSS_DIFF,
        d.GROSS_MTD == '' ? '-' : d.GROSS_MTD == null ? '-' : d.GROSS_MTD,
        d.GROSS_PACE == '' ? '-' : d.GROSS_PACE == null ? '-' : d.GROSS_PACE,
        d.GROSS_TARGET == '' ? '-' : d.GROSS_TARGET == null ? '-' : d.GROSS_TARGET,
        d.GROSS_DIFF == '' ? '-' : d.GROSS_DIFF == null ? '-' : d.GROSS_DIFF,
        d.GROSS_FORECAST == '' ? '-' : d.GROSS_FORECAST == null ? '-' : d.GROSS_FORECAST + '%',
        d.GROSS_LY == '' ? '-' : d.GROSS_LY == null ? '-' : d.GROSS_LY,
        d.LY_GROSS_DIFF == '' ? '-' : d.LY_GROSS_DIFF == null ? '-' : d.LY_GROSS_DIFF,
        d.GP == '' ? '-' : d.GP == null ? '-' : d.GP + '%',

      ];

      const Data1 = worksheet.addRow(obj);
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
          (number > 2 && number <= 6) ||
          (number > 7 && number <= 13) || number == 15 || number == 16
        ) {
          cell.numFmt = '$#,##0';
        }

        if (number != 1) {
          cell.alignment = { vertical: 'top', horizontal: 'center', indent: 1 };
        }
        if (obj[number] < 0) {
          Data1.getCell(number + 1).font = {
            name: 'Arial',
            family: 4,
            size: 9,
            color: { argb: 'FFFF0000' },
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
        for (const d1 of d.Subdata) {
          let obj2 = [
            d1.DEPT == '' ? '-' : d1.DEPT == null ? '-' : d1.DEPT,
            d1.UNITS == '' ? '-' : d1.UNITS == null ? '-' : d1.UNITS,
            d1.SALES_MTD == '' ? '-' : d1.SALES_MTD == null ? '-' : d1.SALES_MTD,
            d1.SALES_PACE == '' ? '-' : d1.SALES_PACE == null ? '-' : d1.SALES_PACE,
            d1.SALES_TARGET == '' ? '-' : d1.SALES_TARGET == null ? '-' : d1.SALES_TARGET,
            d1.SALES_DIFF == '' ? '-' : d1.SALES_DIFF == null ? '-' : d1.SALES_DIFF,
            d1.SALES_FORECAST == '' ? '-' : d1.SALES_FORECAST == null ? '-' : d1.SALES_FORECAST + '%',
            d1.SALES_LY == '' ? '-' : d1.SALES_LY == null ? '-' : d1.SALES_LY,
            d1.LY_GROSS_DIFF == '' ? '-' : d1.LY_GROSS_DIFF == null ? '-' : d1.LY_GROSS_DIFF,
            d1.GROSS_MTD == '' ? '-' : d1.GROSS_MTD == null ? '-' : d1.GROSS_MTD,
            d1.GROSS_PACE == '' ? '-' : d1.GROSS_PACE == null ? '-' : d1.GROSS_PACE,
            d1.GROSS_TARGET == '' ? '-' : d1.GROSS_TARGET == null ? '-' : d1.GROSS_TARGET,
            d1.GROSS_DIFF == '' ? '-' : d1.GROSS_DIFF == null ? '-' : d1.GROSS_DIFF,
            d1.GROSS_FORECAST == '' ? '-' : d1.GROSS_FORECAST == null ? '-' : d1.GROSS_FORECAST + '%',
            d1.GROSS_LY == '' ? '-' : d1.GROSS_LY == null ? '-' : d1.GROSS_LY,
            d1.LY_GROSS_DIFF == '' ? '-' : d1.LY_GROSS_DIFF == null ? '-' : d1.LY_GROSS_DIFF,
            d1.GP == '' ? '-' : d1.GP == null ? '-' : d1.GP + '%',
          ];
          const Data2 = worksheet.addRow(obj2);
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
              (number > 2 && number <= 6) ||
              (number > 7 && number <= 13) || number == 15 || number == 16
            ) {
              cell.numFmt = '$#,##0';
            }
            if (number != 1) {
              cell.alignment = {
                vertical: 'top',
                horizontal: 'center',
                indent: 1,
              };
            }
            if (obj2[number] < 0) {
              Data2.getCell(number + 1).font = {
                name: 'Arial',
                family: 4,
                size: 9,
                color: { argb: 'FFFF0000' },
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
          if (d1.Data2 != undefined) {
            for (const d2 of d1.Data2) {
              let obj3 = [
                d2.ASG_Subtype_Detail == '' ? '-' : d2.ASG_Subtype_Detail == null ? '-' : d2.ASG_Subtype_Detail,
                d2.UNITS == '' ? '-' : d2.UNITS == null ? '-' : d2.UNITS,
                d2.SALES_MTD == '' ? '-' : d2.SALES_MTD == null ? '-' : d2.SALES_MTD,
                d2.SALES_PACE == '' ? '-' : d2.SALES_PACE == null ? '-' : d2.SALES_PACE,
                d2.SALES_TARGET == '' ? '-' : d2.SALES_TARGET == null ? '-' : d2.SALES_TARGET,
                d2.SALES_DIFF == '' ? '-' : d2.SALES_DIFF == null ? '-' : d2.SALES_DIFF,
                d2.SALES_FORECAST == '' ? '-' : d2.SALES_FORECAST == null ? '-' : d2.SALES_FORECAST + '%',
                d2.SALES_LY == '' ? '-' : d2.SALES_LY == null ? '-' : d2.SALES_LY,
                d2.LY_GROSS_DIFF == '' ? '-' : d2.LY_GROSS_DIFF == null ? '-' : d2.LY_GROSS_DIFF,
                d2.GROSS_MTD == '' ? '-' : d2.GROSS_MTD == null ? '-' : d2.GROSS_MTD,
                d2.GROSS_PACE == '' ? '-' : d2.GROSS_PACE == null ? '-' : d2.GROSS_PACE,
                d2.GROSS_TARGET == '' ? '-' : d2.GROSS_TARGET == null ? '-' : d2.GROSS_TARGET,
                d2.GROSS_DIFF == '' ? '-' : d2.GROSS_DIFF == null ? '-' : d2.GROSS_DIFF,
                d2.GROSS_FORECAST == '' ? '-' : d2.GROSS_FORECAST == null ? '-' : d2.GROSS_FORECAST + '%',
                d2.GROSS_LY == '' ? '-' : d2.GROSS_LY == null ? '-' : d2.GROSS_LY,
                d2.LY_GROSS_DIFF == '' ? '-' : d2.LY_GROSS_DIFF == null ? '-' : d2.LY_GROSS_DIFF,
                d2.GP == '' ? '-' : d2.GP == null ? '-' : d2.GP + '%',
              ];
              const Data3 = worksheet.addRow(obj3);
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
                if (
                  (number > 2 && number <= 6) ||
                  (number > 7 && number <= 13) || number == 15 || number == 16
                ) {
                  cell.numFmt = '$#,##0';
                }
                if (number != 1) {
                  cell.alignment = {
                    vertical: 'top',
                    horizontal: 'center',
                    indent: 1,
                  };
                }
                if (obj3[number] < 0) {
                  Data3.getCell(number + 1).font = {
                    name: 'Arial',
                    family: 4,
                    size: 9,
                    color: { argb: 'FFFF0000' },
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
        if (rowIndex > 1 && rowIndex < 19) {
          // Skip the header row
          // Apply conditional alignment based on your conditions
          if (colIndex === 1) {
            // Apply right alignment to the second column
            cell.alignment = {
              horizontal: 'left',
              vertical: 'middle',
              indent: 1,
            };
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
    worksheet.getColumn(8).width = 25;
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
      this.shared.exportToExcel(workbook, 'Parts Summary GL')

    });
    // });
  }




}