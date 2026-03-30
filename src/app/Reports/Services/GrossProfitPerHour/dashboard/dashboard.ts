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
  ServiceData: any = [];
  servicegrossprofitdata: any = [];
  TotalServiceGross: any = [];
  type: any = 'A';
  TotalReport: string = 'T';
  DummyReportType: String = 'T';
  ReportType: any = 'T';
  NoData: boolean = false;
  Paytype: any = ['C', 'W', 'I', 'EW'];
  Subtype: any = [];
  duplicatePaytype: any = ['C', 'W', 'I', 'EW'];
  responcestatus: any = '';
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
    this.shared.setTitle(this.comm.titleName + '-Gross Profit Per Hour');
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
    this.initializeDates('MTD')
    this.setHeaderData()
    this.getSubTypeDetail();
    this.getServiceData();
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
    this.responcestatus = '';
    this.shared.spinner.show();
    this.GetData();
    this.GetTotalData();
  }
  detailHeaders: any = [];
  GetData() {
    console.log(this.detailHeaders, '..................');
    this.DummyReportType = this.ReportType
    this.responcestatus = '';   // ✅ RESET STATUS
    this.servicegrossprofitdata = [];
    this.ServiceData = [];
    const obj = {
      Startdate: this.FromDate,
      Enddate: this.ToDate,
      StoreID: this.storeIds == 0 ? '' : this.storeIds.toString(),
      Paytype: this.ReportType === 'T'
        ? this.Paytype.toString()
        : this.ReportType === 'G'
          ? this.selectedSubType.toString()
          : '',
      type: 'D',
      reporttype: this.ReportType
    };
    this.shared.api.postmethod(
      this.comm.routeEndpoint + 'GetServiceGrossProfits',
      obj
    ).subscribe(
      (res) => {
        if (res.status === 200 && res.response?.length > 0) {
          this.responcestatus += 'I';
          if (this.ReportType === 'T') {
            this.servicegrossprofitdata = res.response;
            this.servicegrossprofitdata.forEach((x: any) => {
              let paytype = ["CP", "WP", "IP", "EWP"];
              paytype.forEach(pt => {
                if (x[pt]) {
                  x[pt] = JSON.parse(x[pt]);
                }
              });
            });
            this.combineIndividualandTotal();
          } else if (this.ReportType === 'G') {
            this.ServiceData = res.response.map((item: any) => {
              const parsedDetails = item.Details
                ? JSON.parse(item.Details)
                : [];
              return {
                ...item,
                Hours: Number(item.Hours || item.hours || 0),
                GrossProfit: Number(item.GrossProfit || 0),
                Discount: Number(item.Discount || 0),
                GP_Hour: Number(item.GP_Hour || 0),
                Details: parsedDetails.map((d: any) => ({
                  ASg_Paytype: d.ASg_Paytype,
                  hourscp: Number(d.hourscp || 0),
                  GrossProfitCP: Number(d.GrossProfitCP || 0)
                }))
              };
            });
            this.detailHeaders = this.selectedSubType.map((name: string) => ({
              ASg_Paytype: name
            }));
            this.ServiceData = this.ServiceData.map((store: any) => {
              const updatedDetails = this.detailHeaders.map((header: any) => {
                const existing = store.Details.find(
                  (d: any) => d.ASg_Paytype === header.ASg_Paytype
                );
                return existing
                  ? existing
                  : {
                    ASg_Paytype: header.ASg_Paytype,
                    hourscp: null,
                    GrossProfitCP: null
                  };
              });
              return {
                ...store,
                Details: updatedDetails
              };
            });
            this.combineIndividualandTotal();
          }
          console.log('Service Data', this.ServiceData);
        } else {
          this.NoData = true;
          this.shared.spinner.hide();
        }
      },
      () => {
        this.toast.show('502 Bad Gate Way Error', 'danger', 'Error');
        this.shared.spinner.hide();
        this.NoData = true;
      }
    );
  }
  GetTotalData() {
    this.TotalServiceGross = [];
    const obj = {
      Startdate: this.FromDate,
      Enddate: this.ToDate,
      StoreID: this.storeIds == 0 ? '' : this.storeIds.toString(),
      Paytype: this.ReportType === 'T'
        ? this.Paytype.toString()
        : this.ReportType === 'G'
          ? this.selectedSubType.toString()
          : '',
      type: 'T',
      reporttype: this.ReportType
    };
    this.shared.api.postmethod(
      this.comm.routeEndpoint + 'GetServiceGrossProfits',
      obj
    ).subscribe(
      (res) => {
        if (res.status === 200 && res.response?.length > 0) {
          this.responcestatus += 'T';
          if (this.ReportType === 'G') {
            const totalItem = res.response[0];
            const parsedDetails = totalItem.Details
              ? JSON.parse(totalItem.Details)
              : [];
            const formattedDetails = parsedDetails.map((d: any) => ({
              ASg_Paytype: d.ASg_Paytype,
              hourscp: Number(d.hourscp || 0),
              GrossProfitCP: Number(d.GrossProfitCP || 0)
            }));
            const orderedDetails = this.selectedSubType.map((name: string) => {
              const existing = formattedDetails.find(
                (d: any) => d.ASg_Paytype === name
              );
              return existing
                ? existing
                : {
                  ASg_Paytype: name,
                  hourscp: 0,
                  GrossProfitCP: 0
                };
            });
            this.TotalServiceGross = [{
              Store_Name: 'REPORTS TOTAL',
              Hours: Number(totalItem.Hours || totalItem.hours || 0),
              GrossProfit: Number(totalItem.GrossProfit || 0),
              Discount: Number(totalItem.Discount || 0),
              GP_Hour: Number(totalItem.GP_Hour || 0),
              Details: orderedDetails
            }];
            console.log('Total Service Gross', this.TotalServiceGross);
          }
          else {
            this.TotalServiceGross = res.response.map((v: any) => ({
              ...v,
              Store_Name: 'REPORTS TOTAL'
            }));
            this.TotalServiceGross.forEach((x: any) => {
              let paytype = ["CP", "WP", "IP", "EWP"];
              paytype.forEach(pt => {
                if (x[pt]) {
                  x[pt] = JSON.parse(x[pt]);
                }
              });
            });
          }
          this.combineIndividualandTotal();
        } else {
          this.NoData = true;
          // this.shared.spinner.hide();
        }
      },
      () => {
        this.toast.show('502 Bad Gate Way Error', 'danger', 'Error');
        // this.shared.spinner.hide();
        this.NoData = true;
      }
    );
  }
  combineIndividualandTotal() {
    if (this.responcestatus === 'IT' || this.responcestatus === 'TI') {
      if (this.ReportType === 'G') {
        if (this.TotalReport === 'B') {
          this.ServiceData = [...this.ServiceData, this.TotalServiceGross[0]];
        } else {
          this.ServiceData = [this.TotalServiceGross[0], ...this.ServiceData];
        }
      } else {
        if (this.TotalReport === 'B') {
          this.servicegrossprofitdata = [
            ...this.servicegrossprofitdata,
            this.TotalServiceGross[0]
          ];
        } else {
          this.servicegrossprofitdata = [
            this.TotalServiceGross[0],
            ...this.servicegrossprofitdata
          ];
        }
        this.ServiceData = this.servicegrossprofitdata;
      }
      this.NoData = false;
    }
    else if (this.responcestatus === 'T') {
      this.ServiceData = this.TotalServiceGross;
      this.NoData = false;
    }
    else if (this.responcestatus === 'I') {
      this.ServiceData = this.ReportType === 'G'
        ? this.ServiceData
        : this.servicegrossprofitdata;
      this.NoData = false;
    }
    else {
      this.NoData = true;
    }
    this.shared.spinner.hide();
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
  setHeaderData() {
    const data = {
      title: 'Gross Profit Per Hour',
      stores: this.storeIds,
      Paytype: this.Paytype.toString(),
      Subtype: this.selectedSubType.toString(),
      ToporBottom: this.TotalReport,
      ReportType: this.ReportType,
      fromdate: this.FromDate,
      todate: this.ToDate,
      groups: this.groupId,
    };
    this.shared.api.SetHeaderData({
      obj: data,
    });
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
  ngAfterViewInit() {
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.comm.pageName == 'Gross Profit Per Hour') {
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
        if (res.obj.title == 'Gross Profit Per Hour') {
          if (res.obj.state == true) {
            this.exportToExcel();
          }
        }
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res) => {
      if (this.print != undefined) {
        if (res.obj.title == 'Gross Profit Per Hour') {
          if (res.obj.statePrint == true) {
            //  this.GetPrintData();
          }
        }
      }
    });
    this.Pdf = this.shared.api.getExportToPDFAllReports().subscribe((res) => {
      if (this.Pdf != undefined) {
        if (res.obj.title == 'Gross Profit Per Hour') {
          if (res.obj.statePDF == true) {
            //   this.generatePDF();
          }
        }
      }
    });
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res) => {
      if (this.email != undefined) {
        if (res.obj.title == 'Gross Profit Per Hour') {
          if (res.obj.stateEmailPdf == true) {
            //   this.sendEmailData(res.obj.Email, res.obj.notes, res.obj.from);
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
  department: any = ['Service', 'Parts'];
  alertState: boolean = false;
  subTypeDetails: any = [];
  selectedSubType: any = [];
  subtypeaction: any = '';
  spinLoader: boolean = false;
  setSubTypeOverallData: any = [];
  mainSubType: any = [];
  singleSubType: any = '';
  spinnerLoader: boolean = false;
  getSubTypeDetail(data?: any) {
    const obj = {
      "AS_IDS": this.storeIds.toString(),
      "DEPTTYPE": 'Service'
    }
    this.spinLoader = true;
    this.subTypeDetails = [];
    this.subtypeaction = 'Y'
    this.spinnerLoader = true;
    this.shared.api.postmethod(this.comm.routeEndpoint + '/GetSubtypeDetailTypes', obj).subscribe((res: any) => {
      if (res.status == 200) {
        this.spinLoader = false;
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
        console.log('Sub Type Details', this.subTypeDetails, this.mainSubType);
        // else{
        //   this.selectedSubType = []
        // }
      } else {
        this.spinnerLoader = false;
      }

    })
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
  multipleorsingle(block: any, e: any, Data?: any) {
    if (block == 'TB') {
      this.TotalReport = e;
    }
    if (block == 'RT') {
      this.ReportType = e;
    }
    if (block == 'PT') {
      const index = this.Paytype.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.Paytype.splice(index, 1);
      } else {
        this.Paytype.push(e);
      }
    }
    if (block === 'ST') {
      // ✅ Ensure arrays are always defined
      this.selectedSubType = this.selectedSubType || [];
      this.mainSubType = this.mainSubType || [];
      const index = this.selectedSubType.findIndex((i: any) => i === e);
      if (index >= 0) {
        this.selectedSubType.splice(index, 1);
        let data = Data.SubData.filter((item: any) =>
          !this.selectedSubType.includes(item.SubtypeParam)
        );
        if (data.length > 0) {
          const mainindex = this.mainSubType.findIndex((i: any) => i === Data.Dept);
          if (mainindex >= 0) {
            this.mainSubType.splice(mainindex, 1);
          }
        }
        if (this.selectedSubType.length === 0 && this.subTypeDetails?.length > 0) {
          this.toast.show('Please Select Atleast One subtype from each Department', 'warning', 'Warning');
          this.alertState = true;
        } else if (data.length === Data.SubData.length) {
          this.toast.show('Please select any one subtype from ' + Data.Dept + ' Department', 'warning', 'Warning');
          this.alertState = true;
        } else {
          this.alertState = false;
        }
        this.allSubTypeSelection();
      } else {
        this.selectedSubType.push(e);
        this.alertState = false;
        let data = Data.SubData.filter((item: any) =>
          !this.selectedSubType.includes(item.SubtypeParam)
        );
        if (data.length === 0) {
          if (!this.mainSubType.includes(Data.Dept)) {
            this.mainSubType.push(Data.Dept);
          }
        }
        this.allSubTypeSelection();
      }
      console.log(this.mainSubType, this.alertState, this.selectedSubType.toString());
    }
  }
  activePopover: number = -1;
  togglePopover(popoverIndex: number) {
    this.activePopover = this.activePopover === popoverIndex ? -1 : popoverIndex;
  }


  viewreport() {
    this.activePopover = -1
    // if (this.selectedDataGrouping.length == 0) {
    //   this.toast.warning('Please Select Atleast One Value from Grouping', '');
    // } else {
    if (this.storeIds.length == 0) {
      this.toast.show('Please Select Atleast One Store', 'warning', 'Warning');
    } else if (this.selectedSubType.length <= 3) {
      this.toast.show('Please select atleast 3 Sub type', 'warning', 'Warning');
    } else {
      if (this.Paytype.length == 0) {
        this.toast.show('Please Select Atleast One Pay Type', 'warning', 'Warning');
      } else {
        this.setHeaderData();
        this.getServiceData()
      }
    }
  }
  ExcelStoreNames: any = [];
  exportToExcel() {
    let storeNames: any = [];
    const store = this.storeIds
    storeNames = this.comm.groupsandstores
      .filter((v: any) => v.sg_id == this.groupId)[0]
      .Stores.filter((item: any) => store.some((cat: any) => cat === item.ID.toString()));
    this.ExcelStoreNames =
      store.length === this.comm.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores.length
        ? 'All Stores'
        : storeNames.map((a: any) => a.storename);
    const ServiceData = this.ServiceData.map((el: any) => Object.assign({}, el));
    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('Gross Profit Per Hour');
    worksheet.views = [{ showGridLines: false }];
    // ============================
    // Report Header
    // ============================
    worksheet.addRow('');
    const titleRow = worksheet.addRow(['Gross Profit Per Hour']);
    titleRow.eachCell((cell: any) => {
      cell.alignment = { indent: 1, vertical: 'top', horizontal: 'left' };
      cell.font = { name: 'Arial', size: 12, bold: true };
    });
    worksheet.mergeCells('A2', 'D2');
    worksheet.addRow('');
    worksheet.addRow([new Date().toLocaleString()]);
    worksheet.addRow(['Report Controls:']).font = { bold: true };
    worksheet.addRow(['Groupings:', 'Groups Name']).font = { bold: true };
    worksheet.addRow(['Timeframe:', `${this.FromDate} to ${this.ToDate}`]).font = { bold: true };
    worksheet.addRow(['Stores:', this.ExcelStoreNames?.toString().replaceAll(',', ', ') || 'All Stores']).font = { bold: true };
    worksheet.addRow(['Filters:', `Pay Type: ${this.ReportType === 'T'
      ? this.Paytype.toString()
      : this.ReportType === 'G'
        ? this.selectedSubType.toString()
        : ''}`]).font = { bold: true };
    worksheet.addRow(['Report Totals:', this.TotalReport === 'T' ? 'Top' : 'Bottom']).font = { bold: true };
    worksheet.addRow('');
    // ============================
    // Table Headers
    // ============================
    const firstHeaderRow = worksheet.addRow([]);
    if (this.ReportType === 'T') {
      firstHeaderRow.getCell(1).value = 'Store Totals';
      worksheet.mergeCells(`A${firstHeaderRow.number}:E${firstHeaderRow.number}`);
      let colIndex = 5;
      this.duplicatePaytype.forEach((type: string) => {
        const headerName =
          type === 'C' ? 'Customer Pay' :
            type === 'W' ? 'Warranty' :
              type === 'EW' ? 'Extended Warranty' :
                type === 'I' ? 'Internal' : type;
        colIndex++; // spacer
        const startCol = colIndex;
        const endCol = colIndex + 3; // merge 4 columns
        worksheet.mergeCells(firstHeaderRow.number, startCol, firstHeaderRow.number, endCol);
        worksheet.getCell(firstHeaderRow.number, startCol).value = headerName;
        worksheet.getCell(firstHeaderRow.number, startCol).alignment = { horizontal: 'center' };
        colIndex = endCol;
      });
    } else if (this.ReportType === 'G') {
      firstHeaderRow.getCell(1).value = 'Store Totals';
      worksheet.mergeCells(`A${firstHeaderRow.number}:E${firstHeaderRow.number}`);
      let colIndex = 5;
      this.detailHeaders.forEach((header: any) => {
        colIndex++; // spacer
        const startCol = colIndex;
        const endCol = colIndex + 1; // 2 columns per header
        worksheet.mergeCells(firstHeaderRow.number, startCol, firstHeaderRow.number, endCol);
        worksheet.getCell(firstHeaderRow.number, startCol).value = header.ASg_Paytype;
        worksheet.getCell(firstHeaderRow.number, startCol).alignment = { horizontal: 'center' };
        colIndex = endCol;
      });
    }
    // Apply style to first header row
    firstHeaderRow.eachCell((cell) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF337AB7' } };
      cell.font = { color: { argb: 'FFFFFFFF' }, bold: true };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      };
    });
    // ============================
    // Second Header Row
    // ============================
    const secondHeaderRow = worksheet.addRow(['Store Name', 'Hours', 'Gross Profit', 'Discount', 'GP per Hour']);
    let colIndex = 5;
    if (this.ReportType === 'T') {
      this.duplicatePaytype.forEach(() => {
        colIndex++; // spacer
        secondHeaderRow.getCell(colIndex).value = 'Hours';
        secondHeaderRow.getCell(colIndex + 1).value = 'Gross Profit';
        secondHeaderRow.getCell(colIndex + 2).value = 'Discount';
        secondHeaderRow.getCell(colIndex + 3).value = 'GP per Hour';
        colIndex += 3;
      });
    } else if (this.ReportType === 'G') {
      this.detailHeaders.forEach(() => {
        colIndex++; // spacer
        secondHeaderRow.getCell(colIndex).value = 'Hours';
        secondHeaderRow.getCell(colIndex + 1).value = 'Gross Profit';
        colIndex += 1;
      });
    }
    // Style second header row
    secondHeaderRow.eachCell((cell) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF5BC0DE' } };
      cell.font = { color: { argb: 'FFFFFFFF' }, bold: true };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      };
    });
    // ============================
    // Table Data
    // ============================
    ServiceData.forEach((sgp: any) => {
      const row: any[] = [
        sgp.Store_Name || '-',
        sgp.Hours || 0,
        sgp.GrossProfit || 0,
        sgp.Discount || 0,
        sgp.GP_Hour || 0,
      ];
      if (this.ReportType === 'T') {
        this.duplicatePaytype.forEach((type: string) => {
          let payData: any;
          if (type === 'C') payData = sgp.CP?.[0];
          if (type === 'W') payData = sgp.WP?.[0];
          if (type === 'EW') payData = sgp.EWP?.[0];
          if (type === 'I') payData = sgp.IP?.[0];
          row.push(
            payData?.HoursCP || 0,
            payData?.GrossProfitCP || 0,
            payData?.DiscountCP || 0,
            payData?.GP_Hourcp || 0
          );
        });
      } else if (this.ReportType === 'G') {
        this.detailHeaders.forEach((_: any, i: any) => {
          row.push(
            sgp.Details?.[i]?.hourscp || 0,
            sgp.Details?.[i]?.GrossProfitCP || 0
          );
        });
      }
      const dataRow = worksheet.addRow(row);
      // Format cells for each row
      dataRow.eachCell((cell, colNumber) => {
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' },
        };
        // Dollar format for Gross Profit & Discount
        const header = secondHeaderRow.getCell(colNumber).value?.toString();
        if (header?.includes('Gross Profit') || header?.includes('Discount')) {
          cell.numFmt = '"$"#,##0';
        } else if (header?.includes('Hours') || header?.includes('GP')) {
          cell.numFmt = '#,##0';
        }
      });
      if (sgp.Store_Name === 'REPORTS TOTAL') dataRow.font = { bold: true };
    });
    // ============================
    // Column widths
    // ============================
    worksheet.columns.forEach((col, index) => {
      col.width = index === 0 ? 30 : 15;
    });
    // ============================
    // Apply Filter
    // ============================
    worksheet.autoFilter = {
      from: { row: secondHeaderRow.number, column: 1 },
      to: { row: secondHeaderRow.number, column: worksheet.columnCount },
    };
    // ============================
    // Save
    // ============================

    this.shared.exportToExcel(workbook, 'GrossProfitPerHour.xlsx');

  }
}