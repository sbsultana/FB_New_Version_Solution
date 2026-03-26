import { Component, OnInit, ViewChild, ElementRef, HostListener } from '@angular/core';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { Setdates } from '../../../../Core/Providers/SetDates/setdates';
import { BsDatepickerConfig, BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { common } from '../../../../common';
import { Subscription } from 'rxjs';
import { Stores } from '../../../../CommonFilters/stores/stores';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
import { CurrencyPipe } from '@angular/common';
import { PartsTrendingGraph } from '../parts-trending-graph/parts-trending-graph';
import { Router } from '@angular/router';
@Component({
  selector: 'app-dashboard',
  imports: [SharedModule, BsDatepickerModule, Stores],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
})
export class Dashboard {
  Current_Date: any;
  NoData: boolean = false;
  date: any = ''
  Month: any;
  PreviousMonths: any = '13';
  Filter: any = 'PartsTrend';
  SourceFilter: any = 'All';
  ReportTotal: any = 'B';
  reportOpenSub!: Subscription;
  reportGetting!: Subscription;
  Pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;
  otherstoreid: any = '';
  selectedotherstoreids: any = '';
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
  month!: Date;
  DuplicatDate!: Date;
  minDate!: Date;
  maxDate!: Date;
  bsConfig: Partial<BsDatepickerConfig> = {
    dateInputFormat: 'MMMM/YYYY',
    minMode: 'month'
  };
  constructor(public shared: Sharedservice, public setdates: Setdates, private comm: common, private cp: CurrencyPipe, private toast: ToastService, private router: Router) {
    let today = new Date();
    let enddate = new Date(today.setDate(today.getDate() - 1));
    this.Month = new Date(enddate)
    this.date = new Date(enddate.setMonth(enddate.getMonth() - 1))
    this.maxDate = new Date();
    this.minDate = new Date();
    this.minDate.setFullYear(this.maxDate.getFullYear() - 3);
    this.maxDate.setMonth(this.maxDate.getMonth());
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
    this.shared.setTitle(this.comm.titleName + '-Parts Trending Report');
    this.setHeaderData()
    this.DataSelection(this.Filter);
  }
  ngOnInit(): void { }
  setHeaderData() {
    const data = {
      title: 'Parts Trending Report',
      Month: this.date,
      stores: this.storeIds,
      filter: this.Filter,
      source: this.SourceFilter,
      reporttotal: this.ReportTotal,
      PreviousMonths: this.PreviousMonths,
      groups: this.groupId,
      otherstoreids: this.otherstoreid, selectedotherstoreids: this.selectedotherstoreids
    };
    this.shared.api.SetHeaderData({
      obj: data,
    });
  }
  EndDate: any;
  dates: any;
  storeName: any;
  PartsTrendingKeys: any = [];
  ServiceTrendingSubKeys: any = [];
  ServiceTrendingXlm: any = [];
  PartsTrendingData: any = [];
  PartsTrendTotalsData: any = [];

  GetDataByMonths() {
    this.dates = [];
    this.PartsTrendingData = [];
    this.shared.spinner.show();
    let currentDate = new Date(this.date);
    for (let i = 0; i < 12; i++) {
      const tempDate = new Date(currentDate);
      tempDate.setMonth(currentDate.getMonth() - i);
      this.dates.push(tempDate);
    }
    let SubStoreArr: any = [];

    const DateToday = this.shared.datePipe.transform(new Date(this.Month), 'yyyy-MM-dd');
    const obj = {
      startdate: DateToday,
      stores: this.storeIds.toString(),
      count: this.PreviousMonths,
      Source: this.SourceFilter == 'All' ? '' : this.SourceFilter == undefined ? '' : this.SourceFilter,
      type: 'D',
    };
    console.log(obj);
    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetPartsTrendDetails', obj).subscribe(
      (x: any) => {
        const currentTitle = document.title;
        if (x.status == 200) {
          this.PartsTrendingData = x.response;
          const serviceKeys = Object.keys(x.response[0]).slice(5);
          this.PartsTrendingKeys = serviceKeys;
          this.PartsTrendingData = x.response.reduce(
            (r: any, { Store_Name }: any) => {
              if (!r.some((o: any) => o.Store_Name == Store_Name)) {
                r.push({
                  Store_Name,
                  sub: x.response.filter(
                    (v: any) => v.Store_Name == Store_Name
                  ),
                });
              }
              return r;
            },
            []
          );
          this.PartsTrendingData.forEach((e: any, i: any) => {
            SubStoreArr.push(
              e.sub.reduce((r: any, { paytype }: any) => {
                // r.Lable1 = e.Lable1;
                if (!r.some((o: any) => o.paytype == paytype)) {
                  // console.log(r);
                  r.push({
                    paytype,
                    subcategory: e.sub.filter((v: any) => v.paytype == paytype),
                  });
                }
                return r;
              }, [])
            );
            e.sub = SubStoreArr[i];
          });
          this.GetPartsTrendTotals();
          this.shared.spinner.hide();
          console.log('PartsTrendingKeys', this.PartsTrendingKeys);
          console.log('Service Trending Data', this.PartsTrendingData);
          this.shared.spinner.hide();
          if (this.PartsTrendingKeys && this.PartsTrendingKeys.length > 0) {
            this.NoData = false;
            this.shared.spinner.hide();
          } else {
            this.NoData = true;
            this.shared.spinner.hide();
          }
        }
      },
      () => { }
    );
  }
  GetPartsTrendTotals() {
    const DateToday = this.shared.datePipe.transform(new Date(this.Month), 'yyyy-MM-dd');
    const obj = {
      startdate: DateToday,
      stores: this.storeIds.toString(),
      count: this.PreviousMonths,
      Source:
        this.SourceFilter == 'All'
          ? ''
          : this.SourceFilter == undefined
            ? ''
            : this.SourceFilter,
      type: 'T',
    };
    console.log(obj);

    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetPartsTrendDetails', obj).subscribe(
      (x: any) => {
        if (x.status == 200) {
          this.PartsTrendTotalsData = x.response;
          // const serviceKeys = Object.keys(x.response[0]).slice(4);
          // this.PartsTrendingKeys = serviceKeys;
          this.PartsTrendTotalsData = x.response.reduce(
            (r: any, { Store_Name }: any) => {
              if (!r.some((o: any) => o.Store_Name == Store_Name)) {
                r.push({
                  Store_Name,
                  sub: x.response.filter(
                    (v: any) => v.Store_Name == Store_Name
                  ),
                });
              }
              return r;
            },
            []
          );
          this.PartsTrendTotalsData.forEach((e: any) => {
            const subCategories = e.sub.reduce((r: any, { paytype }: any) => {
              if (!r.some((o: any) => o.paytype == paytype)) {
                r.push({
                  paytype,
                  subcategory: e.sub.filter((v: any) => v.paytype == paytype),
                });
              }
              return r;
            }, []);
            e.sub = subCategories;
          });

          if (this.ReportTotal == 'B') {
            this.PartsTrendingData.push(this.PartsTrendTotalsData[0]);
          } else if (this.ReportTotal == 'T') {
            this.PartsTrendingData.unshift(this.PartsTrendTotalsData[0]);
          }
          console.log('Service Trend Data', this.PartsTrendingData);
        }
      },
      () => { }
    );
  }
  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    }
    return false;
  }
  SFstate: any;
  StoreCodes: any;
  block: any = '';
  Report: any = '';
  ngAfterViewInit(): void {
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.comm.pageName == 'Parts Trending Report') {
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
        this.SFstate = res.obj.state;
        if (res.obj.title == 'Parts Trending Report') {
          if (res.obj.state == true) {
            this.exportAsXLSX();
          }
        }
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res) => {
      if (this.print != undefined) {
        if (res.obj.title == 'Parts Trending Report') {
          if (res.obj.statePrint == true) {
            // this.GetPrintData();
          }
        }
      }
    });
    this.Pdf = this.shared.api.getExportToPDFAllReports().subscribe((res) => {
      if (this.Pdf != undefined) {
        if (res.obj.title == 'Parts Trending Report') {
          if (res.obj.statePDF == true) {
            // this.generatePDF();
          }
        }
      }
    });
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res) => {
      if (this.email != undefined) {
        if (res.obj.title == 'Parts Trending Report') {
          if (res.obj.stateEmailPdf == true) {
            // this.sendEmailData(res.obj.Email, res.obj.notes, res.obj.from);
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
  DataSelection(Val: any) {
    if (this.Filter == 'PartsTrend') {
      if (this.storeIds != '' || this.selectedotherstoreids != '') {
        this.GetDataByMonths();
      }
    }
  }
  ValueFormat: any;
  openGraph(monthname: any, dates: any, Obj: any, SummaryType: any) {
    console.log(monthname, dates, Obj, SummaryType);
    const PTRgraph = this.shared.ngbmodal.open(PartsTrendingGraph, {
      size: 'xl',
      backdrop: 'static',
    });
    if (Obj.variable == 'GP') {
      this.ValueFormat = 'Percentage';
    }
    if (
      Obj.variable == 'GROSS' ||
      Obj.variable == 'SALE' ||
      Obj.variable == 'ELR'
    ) {
      this.ValueFormat = 'Currancy';
    } else {
      this.ValueFormat = 'Number';
    }
    PTRgraph.componentInstance.PTRgraphdetails = {
      Object: Obj,
      DATES: dates,
      NAME: Obj.Store_Name,
      ValueFormat: this.ValueFormat,
      STORES: this.storeIds,
      SUMMARYTYPE: SummaryType,
    };
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
  singleSelect(e: string) {
    this.PreviousMonths = e;
  }
  navigateToParts(value: string) {
    this.Filter = value;
  }
  PartsSource(value: string) {
    this.SourceFilter = value;
  }
  TotalReport(Val: any) {
    this.ReportTotal = Val
  }
  activePopover: number = -1;
  togglePopover(popoverIndex: number) {
    this.activePopover = this.activePopover === popoverIndex ? -1 : popoverIndex;
  }


  viewreport() {
    this.activePopover = -1
    if (this.Filter === 'PartsTrend') {
      if (this.storeIds.length == 0) {
        this.toast.show('Please select atleast one Store', 'warning', 'Warning');
      } else {
        this.setHeaderData();
        this.DataSelection(this.Filter);

      }
    } else if (this.Filter === 'ServiceTrend') {
      this.router.navigate(['/ServiceTrendingReport']);
    }
  }
  ExcelStoreNames: any = []
  exportAsXLSX(): void {
    let storeNames: any = [];
    let store = this.storeIds.split(',');
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
    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('Parts Trending Report');
    worksheet.views = [
      {
        state: 'frozen',
        ySplit: 13, // Number of rows to freeze (2 means the first two rows are frozen)
        topLeftCell: 'A14', // Specify the cell to start freezing from (in this case, the third row)
        showGridLines: false,
      },
    ];
    worksheet.addRow('');
    const titleRow = worksheet.addRow(['Parts Trending Report']);
    titleRow.eachCell((cell, number) => {
      cell.alignment = { indent: 1, vertical: 'middle', horizontal: 'left' };
    });
    titleRow.font = { name: 'Arial', family: 4, size: 12, bold: true };
    titleRow.worksheet.mergeCells('A2', 'D2');

    worksheet.addRow('');
    const DateToday = this.shared.datePipe.transform(
      new Date(),
      'MM/dd/yyyy h:mm:ss a'
    );
    worksheet.addRow([DateToday]).font = { name: 'Arial', family: 4, size: 9 };
    const PartsTrendingData = this.PartsTrendingData.map((_arrayElement: any) =>
      Object.assign({}, _arrayElement)
    );
    const StartDate = this.shared.datePipe.transform(
      this.PartsTrendingKeys[0],
      'MMM yyyy'
    );
    const lastIndex = this.PartsTrendingKeys.length;
    const EndDate = this.shared.datePipe.transform(
      this.PartsTrendingKeys[this.PartsTrendingKeys.length - 1],
      'MMM yyyy'
    );
    const Header = [StartDate + ' - ' + EndDate];
    for (let i = 0; i < this.PartsTrendingKeys.length; i++) {
      Header.push(this.PartsTrendingKeys[i]);
    }
    const ReportFilter = worksheet.addRow(['Report Filters :']);
    ReportFilter.font = { name: 'Arial', family: 4, size: 10, bold: true };

    const SummaryType = worksheet.addRow(['Summary Type :']);
    const summarytype = worksheet.getCell('B6');
    summarytype.value = 'Month Summary';
    summarytype.font = { name: 'Arial', family: 4, size: 9 };
    summarytype.alignment = { vertical: 'middle', horizontal: 'left' };
    SummaryType.getCell(1).font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
    };
    const DateMonth = worksheet.addRow(['Month :']);
    const datemonth = worksheet.getCell('B7');
    datemonth.value = this.Month;
    datemonth.font = { name: 'Arial', family: 4, size: 9 };
    datemonth.alignment = { vertical: 'middle', horizontal: 'left' };
    DateMonth.getCell(1).font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
    };
    const Groups = worksheet.getCell('A8');
    Groups.value = 'Group :';
    const groups = worksheet.getCell('B8');
    groups.value = this.comm.groupsandstores.filter(
      (val: any) => val.sg_id == this.groupId.toString()
    )[0].sg_name;
    groups.font = { name: 'Arial', family: 4, size: 9 };
    groups.alignment = {
      vertical: 'middle',
      horizontal: 'left',
      wrapText: true,
    };
    worksheet.mergeCells('B9', 'K11');
    const Stores = worksheet.getCell('A9');
    Stores.value = 'Stores :';
    const stores = worksheet.getCell('B9');
    stores.value =
      this.ExcelStoreNames == 0
        ? '-'
        : this.ExcelStoreNames == null
          ? '-'
          : this.ExcelStoreNames.toString().replaceAll(',', ', ');
    stores.font = { name: 'Arial', family: 4, size: 9 };
    stores.alignment = {
      vertical: 'top',
      horizontal: 'left',
      wrapText: true,
    };
    Stores.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
    };
    worksheet.addRow('');
    const headerRow = worksheet.addRow(Header);
    console.log(headerRow);
    headerRow.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
      color: { argb: 'FFFFFF' },
    };
    headerRow.alignment = {
      indent: 1,
      vertical: 'middle',
      horizontal: 'center',
    };
    headerRow.height = 20;
    headerRow.eachCell((cell, number) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '2a91f0' },
        bgColor: { argb: 'FF0000FF' },
      };
      cell.border = { right: { style: 'thin' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });
    for (const d of PartsTrendingData) {
      const StoreNameRow = worksheet.addRow([d.Store_Name]);
      StoreNameRow.alignment = { vertical: 'middle', horizontal: 'center' };
      StoreNameRow.getCell(1).alignment = {
        indent: 1,
        vertical: 'middle',
        horizontal: 'left',
      };
      StoreNameRow.font = { name: 'Arial', family: 4, bold: true, size: 9 };

      if (d.sub != undefined) {
        for (const d1 of d.sub) {
          let formattedPaytype: string;
          switch (d1.paytype) {
            case 'Totals':
              formattedPaytype = 'TOTAL';
              break;
            case 'C':
              formattedPaytype = 'CUSTOMER PAY';
              break;
            case 'T':
              formattedPaytype = 'WARRANTY';
              break;
            case 'W':
              formattedPaytype = 'WHOLESALE';
              break;
            case 'I':
              formattedPaytype = 'INTERNAL';
              break;
            case 'R':
              formattedPaytype = 'RETAIL';
              break;
            case undefined:
            case '':
              formattedPaytype = '--';
              break;
            default:
              formattedPaytype =
                d1.paytype.toString().slice(0, 26) +
                (d1.paytype.toString().length > 26 ? '...' : '');
              break;
          }

          const PaytypeRow = worksheet.addRow([formattedPaytype]);
          PaytypeRow.outlineLevel = 1;
          PaytypeRow.font = {
            name: 'Arial',
            family: 4,
            bold: true,
            size: 9,
            color: { argb: '2a91f0' },
          };
          PaytypeRow.alignment = { vertical: 'middle', horizontal: 'center' };
          PaytypeRow.getCell(1).alignment = {
            indent: 2,
            vertical: 'middle',
            horizontal: 'left',
          };

          for (const d2 of d1.subcategory) {
            const variableValue =
              d2.variable === '' || d2.variable === null
                ? '-'
                : d2.variable === 'GP'
                  ? 'GP%'
                  : d2.variable;
            const row = [variableValue];
            for (const item of this.PartsTrendingKeys) {
              row.push(d2[item] === '' || d2[item] === null ? '-' : d2[item]);
            }

            const Data2 = worksheet.addRow(row);
            Data2.eachCell((cell: any, number: any) => {
              cell.state = {
                state: 'show',
                collapsed: false,
              };
              if (
                !(
                  d2.variable === 'RO COUNT' ||
                  d2.variable === 'HOURS' ||
                  d2.variable === 'HOURS / RO' ||
                  d2.variable === 'VIN COUNT' ||
                  d2.variable === 'GP'
                )
              ) {
                cell.numFmt = '$#,##0';
              } else {
                cell.numFmt = '#,##0';
              }
            });

            Data2.font = { name: 'Arial', family: 4, size: 8 };
            Data2.alignment = {
              vertical: 'middle',
              horizontal: 'right',
              indent: 1,
            };
            Data2.getCell(1).alignment = {
              indent: 3,
              vertical: 'middle',
              horizontal: 'left',
            };
            if (Data2.number % 2) {
              Data2.eachCell((cell: any, number: any) => {
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

    worksheet.eachRow((row, rowIndex) => {
      row.eachCell((cell, colIndex) => {
        if (rowIndex > 1 && rowIndex < 12) {
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
    worksheet.getColumn(8).width = 15;
    worksheet.getColumn(9).width = 15;
    worksheet.getColumn(10).width = 15;
    worksheet.getColumn(11).width = 15;
    worksheet.getColumn(12).width = 15;
    worksheet.getColumn(13).width = 15;
    worksheet.getColumn(14).width = 15;
    worksheet.getColumn(15).width = 15;
    worksheet.addRow([]);

    this.shared.exportToExcel(workbook, 'Parts Trending Report');

  }
}
