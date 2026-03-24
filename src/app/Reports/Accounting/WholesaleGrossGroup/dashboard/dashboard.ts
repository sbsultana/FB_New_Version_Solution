import { Component, Injector, HostListener } from '@angular/core';
import { Api } from '../../../../Core/Providers/Api/api';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { NgbActiveModal, NgbModal } from '@ng-bootstrap/ng-bootstrap';
import { Subscription } from 'rxjs';
import { Workbook } from 'exceljs';
import FileSaver from 'file-saver';
import html2canvas from 'html2canvas';
import jsPDF from 'jspdf';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
const EXCEL_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
const EXCEL_EXTENSION = '.xlsx';
import numeral from 'numeral';

import { BsDatepickerConfig, BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { Stores } from '../../../../CommonFilters/stores/stores';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';

@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [SharedModule, BsDatepickerModule, Stores],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
})


export class Dashboard {
  ReconAnalysisData: any = []
  expanddataarray: any = [];
  NoData: any = ''

  Pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;

  reporttotals: any = 'B'
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
  activePopover: number = -1;

  month!: Date;
  DupMonth: Date = new Date();
  DuplicatDate!: Date;
  minDate!: Date;
  maxDate!: Date;
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
  constructor(public apiSrvc: Api, private ngbmodalActive: NgbActiveModal,
    private toast: ToastService, private injector: Injector, public shared: Sharedservice,) {

    let today = new Date();
    let enddate = new Date(today.setDate(today.getDate() - 1));
    this.month = new Date(enddate.setMonth(enddate.getMonth() - 1))
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
    this.shared.setTitle(this.shared.common.titleName + '-Wholesale Gross Group')

    this.setHeaderData()
    this.GetReconAnalysis('', 0, '', '');
    // this.GetReconAnalysis()
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
    const data = {
      title: 'Wholesale Gross Group',
      stores: this.storeIds,
      groups: this.groupId,
      month: this.month
    };
    this.apiSrvc.SetHeaderData({
      obj: data,
    });
  }


  isDesc: boolean = false;
  column: string = '';
  sort(property: any, data: any, block: any) {
    this.isDesc = !this.isDesc; //change the direction
    this.column = property;
    let direction = this.isDesc ? 1 : -1;
    if (block == 'M') {
      let fakedata = [...data.filter((val: any) => val.Store.toLowerCase() != 'report total')]
      fakedata.sort(function (a: any, b: any) {
        if (a[property] < b[property]) {
          return -1 * direction;
        } else if (a[property] > b[property]) {
          return 1 * direction;
        } else {
          return 0;
        }
      });
      if (this.reporttotals == 'T') {
        this.ReconAnalysisData = [...this.individualData, ...fakedata]
      }
      else {
        this.ReconAnalysisData = [...fakedata, ...this.individualData]
      }
    } else {
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

    block == 'M' ? this.expanddataarray = [] : ''
  }
  individualData: any = []
  GetReconAnalysis(status: any, store: any, account: any, j: any) {

    if (status == '') {
      this.ReconAnalysisData = [];
      this.NoData = '';
      this.shared.spinner.show();
      this.DupMonth = this.month
    }
    const obj = {
      "AS_IDs": status ? store : this.storeIds.toString(),
      "Date": this.shared.datePipe.transform(this.month, 'yyyy-MM-dd'),
      "Report_Level": status ? 'D' : 'S'
    };
    this.apiSrvc.postmethod(this.shared.common.routeEndpoint + 'GetGLWholeSaleGrossGroup', obj).subscribe(
      (res) => {
        if (res.status == 200) {

          if (res.response != undefined) {
            if (res.response.length > 0) {
              console.log(res.response);

              if (status == '') {
                this.NoData = '';
                let totals = res.response.filter((e: any) => e.Store == 'Total').map((v: any) => ({
                  ...v, accountnumber: 'Total'
                }))
                let storesData = res.response.filter((e: any) => e.Store != 'Total')
                let completeData = [...storesData, ...totals].map((v: any) => ({
                  ...v, subdata: []
                }))
                console.log(completeData, 'Complete Data');
                let reporttotal = completeData.filter((e: any) => e.Store == 'Report Total')
                let individual = completeData.filter((e: any) => e.Store != 'Report Total')
                this.individualData = completeData.filter((e: any) => e.Store.toLowerCase() == 'report total')
                if (individual && individual.length > 0) {
                  if (this.reporttotals == 'T') {
                    this.ReconAnalysisData = [...reporttotal, ...individual]
                  }
                  else {
                    this.ReconAnalysisData = [...individual, ...reporttotal]

                  }
                }

                individual && individual.length > 0 ? this.NoData = '' : this.NoData = 'No data Found'
                // this.ReconAnalysisData = data
                console.log(this.ReconAnalysisData);

              } else {
                console.log(j);

                this.ReconAnalysisData[j].subdata = res.response
              }
              console.log(this.ReconAnalysisData);
              this.shared.spinner.hide();
            } else {
              this.shared.spinner.hide();
              this.NoData = 'No Data Found';
            }
          } else {
            this.shared.spinner.hide();
            this.NoData = 'No Data Found';
          }
        } else {
          this.toast.show(res.status, 'danger', 'Error');
          this.shared.spinner.hide();
          this.NoData = 'No Data Found';
        }
      },
      (error) => {
        this.toast.show('502 Bad Gate Way Error', 'danger', 'Error');
        this.shared.spinner.hide();
        this.NoData = 'No data Found';
      }
    );
  }

  getAvg(colname: any) {
    let total: any = 0;
    let reportTotal = this.ReconAnalysisData.filter((val: any) => val.Store == 'Report Total');
    console.log(reportTotal[0][colname], this.ReconAnalysisData.length - 1);
    return (reportTotal[0][colname] / this.ReconAnalysisData.length - 1).toFixed(2)
  }

  viewreport() {
    this.activePopover = -1
    if (this.storeIds.length == 0) {
      this.toast.show('Please select atleast one store', 'warning', 'Warning');
    }

    else {
      // const data = {
      //   Reference: 'Wholesale Gross Group',
      //   storeValues: this.storeIds.toString(),
      //   groups: this.groups.toString(),
      // };
      // this.apiSrvc.SetReports({
      //   obj: data,
      // });
      this.expanddataarray = []

      if (this.storeIds != '') {
        this.GetReconAnalysis('', 0, '', '');
        this.setHeaderData()
      } else {
        this.NoData = '';
      }
    }
  }

  multipleorsingle(block: any, e: any) {
    if (block == 'TB') {
      this.reporttotals = e;
    }
  }

  expandorcollapse(j: any, data: any) {
    const index = this.expanddataarray.findIndex((i: any) => i == j);
    (index >= 0) ? this.expanddataarray.splice(index, 1) : this.expanddataarray.push(j);
    (index >= 0) ? '' : this.GetReconAnalysis('D', data.AS_ID, data.accountnumber, j);

  }

  ngAfterViewInit() {

    this.apiSrvc.getStores().subscribe((res: any) => {
      if (this.shared.common.pageName == 'Wholesale Gross Group') {
        if (res.obj.storesData != undefined) {
          this.groupsArray = res.obj.storesData;
          this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
          this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_name : this.groupName = ''
          this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
          this.getStoresandGroupsValues()
        }
      }
    })

    this.email = this.apiSrvc.getExportToEmailPDFAllReports().subscribe((res) => {
      if (this.email != undefined) {
        if (res.obj.title == 'Wholesale Gross Group') {
          if (res.obj.stateEmailPdf == true) {
            // this.sendEmailData(res.obj.Email, res.obj.notes, res.obj.from);
          }
        }
      }
    });
    this.print = this.apiSrvc.getExportToPrintAllReports().subscribe((res) => {
      if (this.print != undefined) {
        if (res.obj.title == 'Wholesale Gross Group') {
          if (res.obj.statePrint == true) {
            // this.GetPrintData();
          }
        }
      }
    });
    this.Pdf = this.apiSrvc.getExportToPDFAllReports().subscribe((res) => {
      if (this.Pdf != undefined) {
        if (res.obj.title == 'Wholesale Gross Group') {
          if (res.obj.statePDF == true) {
            // this.generatePDF();
          }
        }
      }
    });

    this.excel = this.apiSrvc.getExportToExcelAllReports().subscribe((res) => {
      if (this.excel != undefined) {
        if (res.obj.title == 'Wholesale Gross Group') {
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
  public inTheGreen(value: any): boolean {
    if (value >= 0) {
      return true;
    }
    return false;
  }

  togglePopover(popoverIndex: number) {
    this.activePopover = this.activePopover === popoverIndex ? -1 : popoverIndex;
  }
  getSelectedStoreNames(): string {
    if (!this.storeIds || this.storeIds.length === 0) return '';

    const ids = this.storeIds.toString().split(',');

    const selectedStores = this.stores.filter((s: any) =>
      ids.includes(s.ID.toString())
    );

    return selectedStores.map((s: any) => s.storename).join(', ');
  }
  getReportFilters(): { title: string; filters: any[] } {

    const formattedMonth = this.month
      ? this.shared.datePipe.transform(this.month, 'MMMM yyyy')
      : '';

    return {
      title: 'Wholesale Gross Group',
      filters: [
        {
          label: 'Store',
          value: this.getSelectedStoreNames() || 'All Stores'
        },
        {
          label: 'Group',
          value: this.groupName || ''
        },
        {
          label: 'Month',
          value: formattedMonth || ''
        }
      ]
    };
  }
  addExcelFiltersSection(worksheet: any): number {
    let rowCount = 0;

    const report = this.getReportFilters();

    /*  TITLE (LEFT ALIGNED) */
    const titleRow = worksheet.addRow([report.title]);
    titleRow.font = { bold: true, size: 14 };
    worksheet.mergeCells(`A${rowCount + 1}:G${rowCount + 1}`);
    titleRow.alignment = { horizontal: 'left', vertical: 'middle' };
    rowCount++;

    /* FILTERS */
    report.filters.forEach((filter: any) => {
      const row = worksheet.addRow([`${filter.label}:`, filter.value]);
      row.getCell(1).font = { bold: true };
      rowCount++;
    });

    /* SPACE */
    worksheet.addRow([]);
    rowCount++;

    return rowCount;
  }

  ExcelStoreNames: any = []
  exportToExcel() {
    let storeNames: any[] = [];
    const store = this.storeIds;

    storeNames = this.shared.common.groupsandstores
      .filter((v: any) => v.sg_id == this.groupId)[0]
      .Stores.filter((item: any) => store.includes(item.ID));

    if (
      store.length ==
      this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0]
        .Stores.length
    ) {
      this.ExcelStoreNames = "All Stores";
    } else {
      this.ExcelStoreNames = storeNames.map((a: any) => a.storename);
    }

    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet("Wholesale Gross Group ");

    const DATE_EXTENSION = this.shared.datePipe.transform(new Date(), "MMddyyyy");
    const DateToday = this.shared.datePipe.transform(new Date(), "MM/dd/yyyy h:mm:ss a");

    worksheet.views = [
      { state: "frozen", ySplit: 9, topLeftCell: "A10", showGridLines: false },
    ];

    worksheet.addRow([]);
    const titleRow = worksheet.addRow(["Wholesale Gross Group "]);
    titleRow.font = { bold: true, size: 12 };

    worksheet.addRow([]);
    worksheet.addRow([DateToday]).font = { size: 9 };
    worksheet.addRow(["Selected Details:"]).font = { bold: true, size: 10 };
    worksheet.addRow(["Date:", this.shared.datePipe.transform(this.month, "MMMM yyyy")]);
    worksheet.addRow(["Store:", this.ExcelStoreNames]);
    worksheet.addRow([]);

    const groupHeaderRow = worksheet.addRow([
      '',
      this.shared.datePipe.transform(this.month, 'MMMM yyyy'),
      '',
      '',
      '', ''
    ]);

    groupHeaderRow.font = { bold: true, size: 10 };


    worksheet.mergeCells(`B${groupHeaderRow.number}:G${groupHeaderRow.number}`);


    // worksheet.mergeCells(`E${groupHeaderRow.number}:G${groupHeaderRow.number}`);


    groupHeaderRow.eachCell((cell, colNumber) => {
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'D9D9D9' }
      };
    });
    const headers = ["Store", "Units", "Gross", "PVR", "YTD Units", "YTD Gross", "YTD PVR"];
    const headerRow = worksheet.addRow(headers);

    headerRow.font = { bold: true, color: { argb: "FFFFFFFF" }, size: 9 };
    headerRow.eachCell((cell) => {
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "2a91f0" },
      };
      cell.alignment = { horizontal: "center", vertical: "middle", indent: 1 };
    });

    const columnWidths = new Array(headers.length).fill(10);

    this.ReconAnalysisData.forEach((item: any, index: number) => {
      const rowData = [
        item.Store || "-",
        item.Units ?? "-",
        item.Gross ?? "-",
        item.PVR ?? "-",
        item.YTDUnits ?? "-",
        item.YTDGross ?? "-",
        item.YTDPVR ?? "-"
      ];

      const mainRow = worksheet.addRow(rowData);
      mainRow.font = { size: 9 };

      if (typeof item.Gross === "number") {
        mainRow.getCell(3).numFmt = "$#,##0";
      }
      if (typeof item.PVR === "number") {
        mainRow.getCell(4).numFmt = "$#,##0";
      }
      if (typeof item.YTDGross === "number") {
        mainRow.getCell(6).numFmt = "$#,##0";
      }
      if (typeof item.YTDPVR === "number") {
        mainRow.getCell(7).numFmt = "$#,##0";
      }
      mainRow.eachCell((cell) => {
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "e5e5e5" },
        };
        cell.alignment = { ...cell.alignment, indent: 1 };
      });

      rowData.forEach((val, i) => {
        const length = val?.toString().length || 0;
        columnWidths[i] = Math.max(columnWidths[i], length);
      });

      if (item.Store.toLowerCase() !== "report total") {
        // item.subdata.forEach((sub: any, j: number) => {

        // ✅ FIXED: MATCH HTML EXPANSION LOGIC EXACTLY
        if (item.subdata && this.expanddataarray.includes(index)) {
          const Subheaders = [
            "Accounting Date",
            "Account #", "",
            "Description",
            "Control",
            "Refer",
            "Balance",
          ];

          const subheaderRow = worksheet.addRow(Subheaders);
          subheaderRow.font = { bold: true, color: { argb: "FFFFFFFF" }, size: 9 };
          subheaderRow.eachCell((cell) => {
            cell.fill = {
              type: "pattern",
              pattern: "solid",
              fgColor: { argb: "2a91f0" },
            };
            cell.alignment = {
              horizontal: "center",
              vertical: "middle",
              indent: 1,
            };
          });

          item.subdata.forEach((val: any) => {
            const innerData = [
              val["Accounting Date"]
                ? this.shared.datePipe.transform(val["Accounting Date"], "MM.dd.yyyy")
                : "-",
              val.Account || "-", "",
              val["Account Description"] || "-",
              val.Control ?? "-",
              val.Refer ?? "-",
              val.Balance ?? "-",
            ];

            const subDetailRow = worksheet.addRow(innerData);
            subDetailRow.font = { size: 9 };

            if (typeof val.Balance === "number") {
              subDetailRow.getCell(7).numFmt = "$#,##0";
            }

            subDetailRow.eachCell((cell) => {
              cell.fill = {
                type: "pattern",
                pattern: "solid",
                fgColor: { argb: "d2deed" },
              };
              cell.alignment = { ...cell.alignment, indent: 1 };
            });

            innerData.forEach((val, i) => {
              const length = val?.toString().length || 0;
              columnWidths[i] = Math.max(columnWidths[i], length);
            });
          });
        }
        // });
      }
    });

    const Avg = [
      'Average',
      this.getAvg('Units') ? this.getAvg('Units') : "-",
      this.getAvg('Gross') ? '$' + this.getAvg('Gross') : "-",
      this.getAvg('PVR') ? '$' + this.getAvg('PVR') : "-",
      this.getAvg('YTDUnits') ? this.getAvg('YTDUnits') : "-",
      this.getAvg('YTDGross') ? '$' + this.getAvg('YTDGross') : "-",
      this.getAvg('YTDPVR') ? '$' + this.getAvg('YTDPVR') : "-"
    ];

    const avgRow = worksheet.addRow(Avg);
    avgRow.eachCell((cell) => {
      avgRow.getCell(3).numFmt = "$#,##0";
      avgRow.getCell(4).numFmt = "$#,##0";
      avgRow.getCell(6).numFmt = "$#,##0";
      avgRow.getCell(7).numFmt = "$#,##0";



      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "e5e5e5" },
      };
      cell.alignment = { horizontal: "right", vertical: "middle", indent: 1 };
    });
    avgRow.getCell(1).alignment = { horizontal: "left", vertical: "middle", indent: 1 };


    // avgRow.font = { bold: true, color: { argb: "FFFFFFFF" }, size: 9 };
    // avgRow.eachCell((cell) => {
    //   cell.fill = {
    //     type: "pattern",
    //     pattern: "solid",
    //     fgColor: { argb: "2a91f0" },
    //   };
    //   cell.alignment = { horizontal: "center", vertical: "middle", indent: 1 };
    // });

    columnWidths.forEach((width, i) => {
      worksheet.getColumn(i + 1).width = Math.max(15, width + 2);
    });

    workbook.xlsx.writeBuffer().then((data) => {
      const blob = new Blob([data], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      FileSaver.saveAs(blob, `Wholesale Gross Group_${DATE_EXTENSION}.xlsx`);
    });
  }
}
