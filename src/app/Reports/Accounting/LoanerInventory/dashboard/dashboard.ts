import {
  Component,
  ElementRef,
  Injector,
  Input,
  OnInit,
  SimpleChanges,
  ViewChild,
  HostListener
} from '@angular/core';
import { ActivatedRoute, Router } from '@angular/router';
import { NgxSpinnerService } from 'ngx-spinner';
import { NgbActiveModal, NgbModal } from '@ng-bootstrap/ng-bootstrap';
import { Title } from '@angular/platform-browser';
import { CommonModule, CurrencyPipe, DatePipe, NgStyle } from '@angular/common';
import * as FileSaver from 'file-saver';
import { Workbook } from 'exceljs';
import { jsPDF } from 'jspdf';
import html2canvas from 'html2canvas';
import { Subscription } from 'rxjs';
import { Api } from '../../../../Core/Providers/Api/api';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { BsDatepickerConfig, BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { environment } from '../../../../../environments/environment';
import { common } from '../../../../common';
import { Stores } from '../../../../CommonFilters/stores/stores';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
import { FormsModule } from '@angular/forms';
const EXCEL_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
const EXCEL_EXTENSION = '.xlsx';

@Component({
  selector: 'app-dashboard',
  imports: [CommonModule, SharedModule, BsDatepickerModule, FormsModule, Stores],
  standalone: true,
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
})
export class Dashboard {
  NoData: boolean = false;
  LendersData: any;
  filteredData: any[] = [];
  selectedstrid = 0;
  selectedDate: Date = new Date();
  currentMonth!: Date;
  Month: any = '';
  groups: any = 1;
  StoreName: any;
  selectedstorevalues: any = [];
  gridvisibility: any;
  bsRangeValue!: Date[];
  activePopover: number = -1;
  storeIds: any = '0';
  popup: any = [{ type: 'Popup' }];
  pdfStyleService: any;
  groupsArray: any = [];
  storename: any = ''
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  groupId: any = 0;
  stores: any = [];
  storesFilterData: any = {
    'groupsArray': this.groupsArray, 'groupId': this.groupId, 'storesArray': this.stores, 'storeids': '1', 'type': 'M', 'others': 'N',
    'groupName': this.groupName, 'storename': this.storename, storecount: null, 'storedisplayname': this.storedisplayname
  };
  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent) {
    const clickedInside = (event.target as HTMLElement).closest('.dropdown-toggle, .reportstores-card , .timeframe');
    if (!clickedInside) {
      this.activePopover = -1;
    }
  }
  Pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;
  constructor(
    public Api: Api,
    private spinner: NgxSpinnerService,
    private ngbmodal: NgbModal,
    private ngbmodalActive: NgbActiveModal,
    private toast: ToastService,
    private title: Title,
    private datepipe: DatePipe,
    private comm: common,
    public shared: Sharedservice,
  ) {
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
    this.title.setTitle(this.comm.titleName + '-Loaner Inventory');
    const data = {
      title: 'Loaner Inventory',
      path1: '',
      path2: '',
      path3: '',
      stores: this.storeIds.toString(),
      store: this.storeIds,
      groups: this.groups,
      count: 0,
    };
    this.Api.SetHeaderData({
      obj: data,
    });
    this.GetData();
  }
  togglePopover(popoverIndex: number) {
    this.activePopover = this.activePopover === popoverIndex ? -1 : popoverIndex;
  }
  UnappliedData: any;
  GetData() {
    this.spinner.show()
    this.UnappliedData = [];
    const obj = {
      DealerId: this.storeIds.toString(),
    };
    const curl = environment.apiUrl + this.comm.routeEndpoint + 'GetLoanInventoryV1';
    this.Api
      .postmethod(this.comm.routeEndpoint + 'GetLoanInventoryV1', obj)
      .subscribe(
        (res) => {
          const currentTitle = document.title;
          if (res.status == 200) {
            if (res.response) {
              if (res.response.length > 0) {
                this.UnappliedData = res.response;
                this.spinner.hide();
                this.NoData = false;

              } else {
                this.spinner.hide();
                this.NoData = true;
              }
            } else {
              this.spinner.hide();
              this.NoData = true;
            }
          } else {

            this.toast.show(res.status, 'danger', 'Error');
            this.spinner.hide();
            this.NoData = true;
          }
        },
        (error) => {
          this.toast.show('502 Bad Gate Way Error', 'danger', 'Error');
          this.spinner.hide();
          this.NoData = true;
        }
      );
  }

  applyDateChange() {
    if (!this.storeIds || this.storeIds.length === 0) {
      this.toast.show(
        'Please Select Atleast One Store',
        'warning',
        'Warning'
      );
      return;
    }
    else {
      this.GetData();
    }
  }
  getTotal(columnname: any) {
    let total: any = 0
    this.UnappliedData.some(function (x: any) {
      total += x[columnname]
    })
    return total
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

  sort(property: any) {
    console.log(property);

    this.isDesc = !this.isDesc; //change the direction
    this.column = property;
    let direction = this.isDesc ? 1 : -1;

    this.UnappliedData.sort(function (a: any, b: any) {
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
    this.Api.getStores().subscribe((res: any) => {
      if (this.shared.common.pageName == 'Loaner Inventory') {
        if (res.obj.storesData != undefined) {
          this.groupsArray = res.obj.storesData;
          this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
          this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_name : this.groupName = ''
          this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
          this.getStoresandGroupsValues()
        }
      }
    })
    this.excel = this.Api.getExportToExcelAllReports().subscribe((res) => {
      if (this.excel != undefined) {
        if (res.obj.title == 'Loaner Inventory') {
          if (res.obj.state == true) {
            this.exportToExcel();
          }
        }
      }
    });
    this.email = this.Api.getExportToEmailPDFAllReports().subscribe((res) => {
      if (this.email != undefined) {
        if (res.obj.title == 'Loaner Inventory') {
          if (res.obj.stateEmailPdf == true) {
            this.sendEmailData(res.obj.Email, res.obj.notes, res.obj.from);
          }
        }
      }
    });
    this.print = this.Api.getExportToPrintAllReports().subscribe((res) => {
      if (this.print != undefined) {
        if (res.obj.title == 'Loaner Inventory') {
          if (res.obj.statePrint == true) {
            this.GetPrintData();
          }
        }
      }
    });
    this.Pdf = this.Api.getExportToPDFAllReports().subscribe((res) => {
      if (this.Pdf != undefined) {
        if (res.obj.title == 'Loaner Inventory') {
          if (res.obj.statePDF == true) {
            this.generatePDF();
          }
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

    // this.setHeaderData();
    // this.GetData();

  }
  StoresData(data: any) {
    this.storeIds = data.storeids;
    this.groupId = data.groupId;
    this.storename = data.storename;
    this.groupName = data.groupName;
    this.storecount = data.storecount;
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
  updateVerticalScroll(event: any): void {
    const scrollDemo = document.querySelector('#scrollcent');
    this.Scrollpercent = Math.round(
      (event.target.scrollTop /
        (event.target.scrollHeight - scrollDemo!.clientHeight)) *
      100
    );
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
    return {
      title: 'Loaner Inventory',
      filters: [
        {
          label: 'Store',
          value: this.getSelectedStoreNames() || 'All Stores'
        },
        {
          label: 'Group',
          value: this.groupName || ''
        },
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
  ExcelStoreNames: any = [];
  exportToExcel() {
    let storeNames: any = [];
    let store = this.storeIds.split(',');

    storeNames = this.comm.groupsandstores
      .filter((v: any) => v.sg_id == this.groups)[0]
      .Stores.filter((item: any) =>
        store.some((cat: any) => cat === item.ID.toString())
      );
    if (
      store.length ==
      this.comm.groupsandstores.filter((v: any) => v.sg_id == this.groups)[0]
        .Stores.length
    ) {
      this.ExcelStoreNames = 'All Stores';
    } else {
      this.ExcelStoreNames = storeNames.map(function (a: any) {
        return a.storename;
      });
    }
    const unapplied = this.UnappliedData.map((_arrayElement: any) =>
      Object.assign({}, _arrayElement)
    );
    console.log(unapplied);

    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet('Loaner Inventory');
    worksheet.views = [
      {
        state: 'frozen',    // Enables the frozen pane view
        ySplit: 15,         // Freezes the first 15 rows
        topLeftCell: 'A16', // Sets the top-left cell to start scrolling from after the frozen pane
        showGridLines: false, // Hides grid lines in the worksheet
      },
    ];

    worksheet.addRow('');
    const titleRow = worksheet.addRow(['Loaner Inventory']);
    titleRow.eachCell((cell, number) => {
      cell.alignment = { indent: 1, vertical: 'middle', horizontal: 'left' };
    });
    titleRow.font = { name: 'Arial', family: 4, size: 12, bold: true };
    worksheet.addRow('');

    const DateToday = this.datepipe.transform(
      new Date(),
      'MM/dd/yyyy h:mm:ss a'
    );
    const DATE_EXTENSION = this.datepipe.transform(new Date(), 'MMddyyyy');
    worksheet.addRow([DateToday]).font = { name: 'Arial', family: 4, size: 9 };

    const ReportFilter = worksheet.addRow(['Report Controls :']);
    ReportFilter.font = { name: 'Arial', family: 4, size: 10, bold: true };

    const Groups = worksheet.getCell('A6');
    Groups.value = 'Group :';
    const groups = worksheet.getCell('B6');
    groups.value = this.comm.groupsandstores.filter(
      (val: any) => val.sg_id == this.groups.toString()
    )[0].sg_name;
    groups.font = { name: 'Arial', family: 4, size: 9 };
    groups.alignment = {
      vertical: 'middle',
      horizontal: 'left',
      wrapText: true,
    };

    worksheet.mergeCells('B7', 'K9');
    const Stores = worksheet.getCell('A7');
    Stores.value = 'Stores :';
    const stores = worksheet.getCell('B7');
    stores.value =
      this.ExcelStoreNames == 0
        ? 'All Stores'
        : this.ExcelStoreNames == null
          ? '-'
          : this.ExcelStoreNames.toString().replaceAll(',', ', ');
    stores.font = { name: 'Arial', family: 4, size: 9 };
    stores.alignment = { vertical: 'top', horizontal: 'left', wrapText: true };




    let Headings = [
      'Stock #',
      'Age',
      'Balance',
      'Balance2',
      'Year',
      'Make',
      'Model',
      'Store',

    ];
    const headerRow = worksheet.addRow(Headings);
    headerRow.font = {
      name: 'Arial',
      family: 4,
      size: 8,
      bold: true,
      color: { argb: 'FFFFFF' },
    };
    headerRow.alignment = {
      indent: 1,
      vertical: 'middle',
      horizontal: 'center',
    };
    headerRow.eachCell((cell, number) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '788494' },
        bgColor: { argb: 'FF0000FF' },
      };
      cell.border = { right: { style: 'thin' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });

    for (const d of unapplied) {

      const Data1 = worksheet.addRow([
        d.StockNo == '' ? '-' : d.StockNo == null ? '-' : d.StockNo,
        d.Age == '' ? '-' : d.Age == null ? '-' : d.Age.toString(),
        d.Balance == '' ? '-' : d.Balance == null ? '-' : d.Balance,
        d['Balance2'] == '' ? '-' : d['Balance2'] == null ? '-' : d['Balance2'],
        d.store == '' ? '-' : d.store == null ? '-' : d.store,
        d.Year == '' ? '-' : d.Year == null ? '-' : d.Year,
        d.Make == '' ? '-' : d.Make == null ? '-' : d.Make,
        d.Model == '' ? '-' : d.Model == null ? '-' : d.Model,

      ]);
      // Data1.outlineLevel = 1; // Grouping level 1
      Data1.font = { name: 'Arial', family: 4, size: 8 };
      Data1.alignment = { vertical: 'middle', horizontal: 'center' };
      // Data1.getCell(1).alignment = {indent: 1,vertical: 'middle', horizontal: 'center'}
      Data1.eachCell((cell, number) => {
        cell.border = { right: { style: 'thin' } };
      });
      Data1.eachCell((cell, number) => {
        cell.numFmt = '$#,##0.00';
        if (number == 4 || number == 3) {
          cell.alignment = {
            indent: 1,
            vertical: 'middle',
            horizontal: 'right',
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


    }

    worksheet.getColumn(1).width = 10;
    worksheet.getColumn(1).alignment = {
      indent: 1,
      vertical: 'middle',
      horizontal: 'left',
    };
    worksheet.getColumn(2).width = 25;
    worksheet.getColumn(2).alignment = {
      indent: 1,
      vertical: 'middle',
      horizontal: 'left',
      wrapText: true,
    };
    worksheet.getColumn(3).width = 15;
    worksheet.getColumn(4).width = 13;
    worksheet.getColumn(5).width = 13;
    worksheet.getColumn(6).width = 13;
    // worksheet.getColumn(7).width = 13;
    worksheet.getColumn(7).width = 25;
    worksheet.getColumn(8).width = 10;
    worksheet.getColumn(9).width = 25;
    worksheet.getColumn(10).width = 25;
    worksheet.getColumn(11).width = 10;

    worksheet.addRow([]);
    workbook.xlsx.writeBuffer().then((data: any) => {
      const blob = new Blob([data], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      FileSaver.saveAs(blob, 'Loaner Inventory_' + DATE_EXTENSION + EXCEL_EXTENSION);
    });
    // });
  }


  GetPrintData() {
    window.print();
  }

  generatePDF() {
    this.spinner.show();
    const printContents = document.getElementById('LoanerInventory')!.innerHTML;

    const iframe = document.createElement('iframe');
    iframe.style.position = 'absolute';
    iframe.style.width = '0px';
    iframe.style.height = '0px';
    iframe.style.border = 'none';
    document.body.appendChild(iframe);

    const doc = iframe.contentWindow?.document;
    if (!doc) {
      console.error('Failed to create iframe document');
      return;
    }

    doc.open();
    doc.write(`
            <html>
                <head>
                <title>CIT/Floor Plan</title>
                <style>
                @font-face {
                 font-family: 'GothamBookRegular';
                 src: url('assets/fonts/Gotham\ Book\ Regular.otf') format('otf'), /* Chrome 6+, Firefox 3.6+, IE 9+, Safari 5.1+ */
                      url('assets/fonts/Gotham\ Book\ Regular.otf') format('opentype'); /* Chrome 4+, Firefox 3.5, Opera 10+, Safari 3—5 */
               }
               @font-face {
                 font-family: 'Roboto';
                 src: url('assets/fonts/Roboto-Regular.ttf') format('ttf'), /* Chrome 6+, Firefox 3.6+, IE 9+, Safari 5.1+ */
                      url('assets/fonts/Roboto-Regular.ttf') format('truetype'); /* Chrome 4+, Firefox 3.5, Opera 10+, Safari 3—5 */
               }
               @font-face {
                 font-family: 'RobotoBold';
                 src: url('assets/fonts/Roboto-Bold.ttf') format('ttf'), /* Chrome 6+, Firefox 3.6+, IE 9+, Safari 5.1+ */
                      url('assets/fonts/Roboto-Bold.ttf') format('truetype'); /* Chrome 4+, Firefox 3.5, Opera 10+, Safari 3—5 */
               }
               .d-flex {
                 display: flex !important;
                }
                .justify-content-between {
                 justify-content: space-between !important;
                 }
                .negative {
                 color: red;
                }
                 .time_stamp {
                   // margin-left: -6rem;
                      padding: 0.2rem .5rem;
                       background-color: #fff;
                       color: #2d91f1;
                     border-radius: 6px;
                     width: 7%;
             }
                       .bdr{
                       border:1px solid white !important
                       }
                       .comment {
                       width: fit-content;
                       background-color: #d7d7d7;
                       border-radius: 0 15px 0 15px;
                       padding: .2rem .5rem;
                       margin-left: 2rem;
                       }
                       .text-end {
                        text-align: right !important;
                    }
                    .col-4 {
                      flex: 0 0 auto;
                      width: 33.33333333%;
                  }
                      .divBold {
                        color: #fff;
                        text-transform: capitalize;
                        white-space: nowrap;
                        font-family: "FaktPro-Normal";
                        font-size: 0.9rem;
                    }
                   
                      .cit-scorecard .table>:not(:first-child) {
                        border-top: 0px solid #ffa51a
                      }
                      .printalignment{
                        margin-left:160rem
                        }
                        .print{
                          margin-left:8rem
                        }
                      .cit-scorecard .table {
                        text-align: center;
                        text-transform: capitalize;
                        border: transparent;
                    
                        width: 52%;
                      }
                      .cit-scorecard .table th,
                      .cit-scorecard .table td{
                        white-space: nowrap;
                        vertical-align: top;
                      }
                      .cit-scorecard .table th:first-child {
                        position: sticky;
                        left: 0;
                        z-index: 1;
                        background-color: #337ab7;
                      }
                      .cit-scorecard .table td:first-child {
                        position: sticky;
                        left: 0;
                        z-index: 1;
                        // background-color: #337ab7;
                      }
                      .cit-scorecard .table  tr:nth-child(even)  td:first-child,
                      .cit-scorecard .table  tr:nth-child(even)  td:nth-child(2) {
                        background-color: #ffffff;
                      }
                      .cit-scorecard .table tr:nth-child(odd) {
                        // background-color: #e9ecef;
                        background-color: #ffffff;
                  
                      }
                      .cit-scorecard .table  tr:nth-child(even) {
                        background-color: #ffffff;
                      }
                      .cit-scorecard .table .spacer {
                        // width: 50px !important;
                        background-color: #cfd6de !important;
                        border-left: 1px solid #cfd6de !important;
                        border-bottom: 1px solid #cfd6de !important;
                        border-top: 1px solid #cfd6de !important;
                      }
                      .cit-scorecard .table .hidden {
                        display: none !important;
                      }
                      .cit-scorecard .table .bdr-rt {
                        border-right: 1px solid #abd0ec;
                      }
                      .cit-scorecard .table  thead {
                        position: sticky;
                        top: 0;
                        z-index: 99;
                        font-family: 'FaktPro-Bold';
                        font-size: 0.8rem;
                      }
                      .cit-scorecard .table  thead  th {
                        padding: 5px 10px;
                        margin: 0px;
                      }
                      .cit-scorecard .table  thead .bdr-btm {
                        border-bottom: #005fa3;
                      }
                      .cit-scorecard .table  thead tr:nth-child(1) {
                        background-color: #fff !important;
                        color: #000;
                        // color: #fff;
                        text-transform: uppercase;
                        border-bottom: #cfd6de;
                      }
                      .cit-scorecard .table  thead tr:nth-child(2) {
                        background-color: #337ab7 !important;
                        color: #fff;
                        text-transform: uppercase;
                        border-bottom: #cfd6de;
                        box-shadow: inset 0 1px 0 0 #cfd6de;
                      }
                      .cit-scorecard .table  thead  tr:nth-child(3) {
          
                        background-color: #fff !important;
                        color: #000;
                        text-transform: uppercase;
                        border-bottom: #cfd6de;
                        box-shadow: inset 0 1px 0 0 #cfd6de;
                      }
                      .cit-scorecard .table  thead  tr:nth-child(3) th :nth-child(1) {
                        background-color: #337ab7 !important;
                        color: #fff;
                      }
                      .cit-scorecard .table tbody {
                        font-family: 'FaktPro-Normal';
                        font-size: .9rem;
                      }
                      .cit-scorecard .table tbody td {
                        padding: 2px 10px;
                        margin: 0px;
                        border: 1px solid #cfd6de
                      }
                      .cit-scorecard .table tbody  tr {
                        border-bottom: 1px solid #37a6f8;
                        border-left: 1px solid #37a6f8
                      }
                      .cit-scorecard .table tbody td:first-child {
                        text-align: start;
                        box-shadow: inset -1px 0 0 0 #cfd6de;
                      }
                      .cit-scorecard .table tbody  tr td:not(:first-child) {
                        text-align: right !important;
                
                      }
                      .cit-scorecard .table tbody .sub-title {
                        font-size: .8rem !important;
                      }
                      cit-scorecard .table tbody  .sub-subtitle {
                        font-size: .7rem !important;
                      }
                      cit-scorecard .table tbody  td:nth-child(2) {
                        padding: 2px 10px;
                        margin: 0px;
                      }
                      cit-scorecard .table tbody .text-bold {
                        font-family: 'FaktPro-Bold';
                      }
                      cit-scorecard .table tbody .darkred-bg {
                        background-color: #282828 !important;
                        color: #fff;
                      }
                      cit-scorecard .table tbody .lightblue-bg {
                        background-color: #646e7a !important;
                        color: #fff;
                      }
                      cit-scorecard .table tbody .gold-bg {
                        background-color: #ffa51a;
                        color: #fff;
                      }
                    
                .performance-scorecard .table > :not(:first-child) {
                 border-top: 0px solid #ffa51a;
                }
                .performance-scorecard .table {
                 text-align: center;
                 text-transform: capitalize;
                 border: transparent;
             
                 width: 100%;
                }
                .performance-scorecard .table th,
                .performance-scorecard .table td {
                 white-space: nowrap;
                 vertical-align: top;
               }
               .performance-scorecard .table  th:first-child,
               .performance-scorecard .table td:first-child {
                 // position: sticky;
                 left: 0;
                 z-index: 1;
                 background-color: #337ab7;
               }
               .performance-scorecard .table  tr:nth-child(odd) td:first-child,
               .performance-scorecard .table  tr:nth-child(odd) td:nth-child(2) {
                 background-color: #ffffff;
               }
               .performance-scorecard .table tr:nth-child(even) td:first-child,
               .performance-scorecard .table tr:nth-child(even) td:nth-child(2) {
                 background-color: #ffffff;
               }
               .performance-scorecard .table tr:nth-child(odd) {
                 background-color: #ffffff;
               }
               .performance-scorecard .table tr:nth-child(even) {
                 background-color: #ffffff;
               }
               .performance-scorecard .table .spacer {
                 // width: 50px !important;
                 background-color: #cfd6de !important;
                 border-left: 1px solid #cfd6de !important;
                 border-bottom: 1px solid #cfd6de !important;
                 border-top: 1px solid #cfd6de !important;
               }
               .performance-scorecard .table .bdr-rt {
                 border-right: 1px solid #abd0ec;
               }
               .performance-scorecard .table thead {
                 position: sticky;
                 top: 0;
                 z-index: 99;
                 font-family: "FaktPro-Bold";
                 font-size: 0.8rem;
               }
               .performance-scorecard .table thead th {
                 padding: 5px 10px;
                 margin: 0px;
               }
               .performance-scorecard .table thead .bdr-btm {
                 border-bottom: #005fa3;
               }
               .performance-scorecard .table thead tr:nth-child(1) {
                 background-color: #fff !important;
                 color: #000;
                 text-transform: uppercase;
                 border-bottom: #cfd6de;
               }
               .performance-scorecard .table thead tr:nth-child(2) {
                 background-color: #337ab7 !important;
                 color: #fff;
                 text-transform: uppercase;
                 border-bottom: #cfd6de;
                 box-shadow: inset 0 1px 0 0 #cfd6de;
               }
               .performance-scorecard .table thead tr:nth-child(3) {
                 background-color: #337ab7 !important;
                 color: #fff;
                 text-transform: uppercase;
                 border-bottom: #cfd6de;
                 box-shadow: inset 0 1px 0 0 #cfd6de;
               }
               .performance-scorecard .table tbody {
                 font-family: "FaktPro-Normal";
                 font-size: 0.9rem;
               }
               .performance-scorecard .table tbody  td {
                 padding: 2px 10px;
                 margin: 0px;
                 border: 1px solid #cfd6de;
               }
               .performance-scorecard .table tbody tr {
                 border-bottom: 1px solid #37a6f8;
                 border-left: 1px solid #37a6f8;
               }
               .performance-scorecard .table tbody td:first-child {
                 text-align: start;
                 box-shadow: inset -1px 0 0 0 #cfd6de;
               }
               .performance-scorecard .table tbody .text-bold {
                 font-family: "FaktPro-Bold";
               }
               .performance-scorecard .table tbody .darkred-bg {
                 background-color: #282828 !important;
                 color: #fff;
               }
               .performance-scorecard .table tbody .lightblue-bg {
                 background-color: #646e7a !important;
                 color: #fff;
               }
               .performance-scorecard .table tbody .gold-bg {
                 background-color: #ffa51a;
                 color: #fff;
               }
           
                </style>
                </head>
                <body id='content'>
                    ${printContents}
                </body>
            </html>`);
    doc.close();

    const div = doc.getElementById('content');
    const options = {
      logging: false,
      allowTaint: false,
      useCORS: true,
    };
    if (!div) {
      return;
    }

    // Start generating PDF with limited time
    html2canvas(div, options).then((canvas) => {
      const imgWidth = 285;
      const pageHeight = 204;
      const imgHeight = (canvas.height * imgWidth) / canvas.width;
      const contentDataURL = canvas.toDataURL('image/jpeg', 0.7); // Reduced quality
      const pdfData = new jsPDF('l', 'mm', 'a4', true);
      let position = 5;

      pdfData.addImage(contentDataURL, 'JPEG', 5, position, imgWidth, imgHeight, undefined, 'FAST');

      let heightLeft = imgHeight - pageHeight;
      while (heightLeft >= 0) {
        position = heightLeft - imgHeight;
        pdfData.addPage();
        pdfData.addImage(contentDataURL, 'JPEG', 5, position, imgWidth, imgHeight, undefined, 'FAST');
        heightLeft -= pageHeight;
      }

      pdfData.save('LoanerInventory.pdf');
      this.spinner.hide();
    }).catch(error => {
      console.error('Error generating PDF:', error);
      this.spinner.hide();
    });
  }

  sendEmailData(Email: any, notes: any, from: any) {


    this.spinner.show();
    const printContents = document.getElementById('LoanerInventory')!.innerHTML;
    const iframe = document.createElement('iframe');

    // Make the iframe invisible
    iframe.style.position = 'absolute';
    iframe.style.width = '0px';
    iframe.style.height = '0px';
    iframe.style.border = 'none';

    document.body.appendChild(iframe);

    const doc = iframe.contentWindow?.document;
    if (!doc) {
      console.error('Failed to create iframe document');
      return;
    }

    doc.open();
    doc.write(`
            <html>
                <head>
                <title>Inventory Open RO</title>
                <title>Loaner Inventory</title>
                <style>
                 @font-face {
                  font-family: 'GothamBookRegular';
                  src: url('assets/fonts/Gotham\ Book\ Regular.otf') format('otf'), /* Chrome 6+, Firefox 3.6+, IE 9+, Safari 5.1+ */
                       url('assets/fonts/Gotham\ Book\ Regular.otf') format('opentype'); /* Chrome 4+, Firefox 3.5, Opera 10+, Safari 3—5 */
                }
                @font-face {
                  font-family: 'Roboto';
                  src: url('assets/fonts/Roboto-Regular.ttf') format('ttf'), /* Chrome 6+, Firefox 3.6+, IE 9+, Safari 5.1+ */
                       url('assets/fonts/Roboto-Regular.ttf') format('truetype'); /* Chrome 4+, Firefox 3.5, Opera 10+, Safari 3—5 */
                }
                @font-face {
                  font-family: 'RobotoBold';
                  src: url('assets/fonts/Roboto-Bold.ttf') format('ttf'), /* Chrome 6+, Firefox 3.6+, IE 9+, Safari 5.1+ */
                       url('assets/fonts/Roboto-Bold.ttf') format('truetype'); /* Chrome 4+, Firefox 3.5, Opera 10+, Safari 3—5 */
                }
                .d-flex {
                  display: flex !important;
                 }
                 .justify-content-between {
                  justify-content: space-between !important;
                  }
                 .negative {
                  color: red;
                 }
                    .time_stamp {
    // margin-left: -6rem;
    padding: 0.2rem .5rem;
    background-color: #fff;
    color: #2d91f1;
    border-radius: 6px;
    width: 7%;
    }
    .bdr{
    border:1px solid white !important
    }
     .comment {
          width: fit-content;
          background-color: #d7d7d7;
          border-radius: 0 15px 0 15px;
          padding: .2rem .5rem;
          margin-left: 2rem;
        }
        .text-end {
          text-align: right !important;
      }
      .col-4 {
        flex: 0 0 auto;
        width: 33.33333333%;
    }
        .divBold {
          color: #fff;
          text-transform: capitalize;
          white-space: nowrap;
          font-family: "FaktPro-Normal";
          font-size: 0.9rem;
      }
      .printalignment{
        margin-left:160rem
        }
        .print{
          margin-left:8rem
        }
        .cit-scorecard .table>:not(:first-child) {
          border-top: 0px solid #ffa51a
        }

        .cit-scorecard .table {
          text-align: center;
          text-transform: capitalize;
          border: transparent;
      
          width: 52%;
        }
        .cit-scorecard .table th,
        .cit-scorecard .table td{
          white-space: nowrap;
          vertical-align: top;
        }
        .cit-scorecard .table th:first-child {
          position: sticky;
          left: 0;
          z-index: 1;
          background-color: #337ab7;
        }
        .cit-scorecard .table td:first-child {
          position: sticky;
          left: 0;
          z-index: 1;
          // background-color: #337ab7;
        }
        .cit-scorecard .table  tr:nth-child(even)  td:first-child,
        .cit-scorecard .table  tr:nth-child(even)  td:nth-child(2) {
          background-color: #ffffff;
        }
        .cit-scorecard .table tr:nth-child(odd) {
          // background-color: #e9ecef;
          background-color: #ffffff;
    
        }
        .cit-scorecard .table  tr:nth-child(even) {
          background-color: #ffffff;
        }
        .cit-scorecard .table .spacer {
          // width: 50px !important;
          background-color: #cfd6de !important;
          border-left: 1px solid #cfd6de !important;
          border-bottom: 1px solid #cfd6de !important;
          border-top: 1px solid #cfd6de !important;
        }
        .cit-scorecard .table .hidden {
          display: none !important;
        }
        .cit-scorecard .table .bdr-rt {
          border-right: 1px solid #abd0ec;
        }
        .cit-scorecard .table  thead {
          position: sticky;
          top: 0;
          z-index: 99;
          font-family: 'FaktPro-Bold';
          font-size: 0.8rem;
        }
        .cit-scorecard .table  thead  th {
          padding: 5px 10px;
          margin: 0px;
        }
        .cit-scorecard .table  thead .bdr-btm {
          border-bottom: #005fa3;
        }
        .cit-scorecard .table  thead tr:nth-child(1) {
          background-color: #fff !important;
          color: #000;
          // color: #fff;
          text-transform: uppercase;
          border-bottom: #cfd6de;
        }
        .cit-scorecard .table  thead tr:nth-child(2) {
          background-color: #337ab7 !important;
          color: #fff;
          text-transform: uppercase;
          border-bottom: #cfd6de;
          box-shadow: inset 0 1px 0 0 #cfd6de;
        }
        .cit-scorecard .table  thead  tr:nth-child(3) {

          background-color: #fff !important;
          color: #000;
          text-transform: uppercase;
          border-bottom: #cfd6de;
          box-shadow: inset 0 1px 0 0 #cfd6de;
        }
        .cit-scorecard .table  thead  tr:nth-child(3) th :nth-child(1) {
          background-color: #337ab7 !important;
          color: #fff;
        }
        .cit-scorecard .table tbody {
          font-family: 'FaktPro-Normal';
          font-size: .9rem;
        }
        .cit-scorecard .table tbody td {
          padding: 2px 10px;
          margin: 0px;
          border: 1px solid #cfd6de
        }
        .cit-scorecard .table tbody  tr {
          border-bottom: 1px solid #37a6f8;
          border-left: 1px solid #37a6f8
        }
        .cit-scorecard .table tbody td:first-child {
          text-align: start;
          box-shadow: inset -1px 0 0 0 #cfd6de;
        }
        .cit-scorecard .table tbody  tr td:not(:first-child) {
          text-align: right !important;
  
        }
        .cit-scorecard .table tbody .sub-title {
          font-size: .8rem !important;
        }
        cit-scorecard .table tbody  .sub-subtitle {
          font-size: .7rem !important;
        }
        cit-scorecard .table tbody  td:nth-child(2) {
          padding: 2px 10px;
          margin: 0px;
        }
        cit-scorecard .table tbody .text-bold {
          font-family: 'FaktPro-Bold';
        }
        cit-scorecard .table tbody .darkred-bg {
          background-color: #282828 !important;
          color: #fff;
        }
        cit-scorecard .table tbody .lightblue-bg {
          background-color: #646e7a !important;
          color: #fff;
        }
        cit-scorecard .table tbody .gold-bg {
          background-color: #ffa51a;
          color: #fff;
        }
      
                 .performance-scorecard .table > :not(:first-child) {
                  border-top: 0px solid #ffa51a;
                 }
                 .performance-scorecard .table {
                  text-align: center;
                  text-transform: capitalize;
                  border: transparent;
              
                  width: 100%;
                 }
                 .performance-scorecard .table th,
                 .performance-scorecard .table td {
                  white-space: nowrap;
                  vertical-align: top;
                }
                .performance-scorecard .table  th:first-child,
                .performance-scorecard .table td:first-child {
                  // position: sticky;
                  left: 0;
                  z-index: 1;
                  background-color: #337ab7;
                }
                .performance-scorecard .table  tr:nth-child(odd) td:first-child,
                .performance-scorecard .table  tr:nth-child(odd) td:nth-child(2) {
                  background-color: #ffffff;
                }
                .performance-scorecard .table tr:nth-child(even) td:first-child,
                .performance-scorecard .table tr:nth-child(even) td:nth-child(2) {
                  background-color: #ffffff;
                }
                .performance-scorecard .table tr:nth-child(odd) {
                  background-color: #ffffff;
                }
                .performance-scorecard .table tr:nth-child(even) {
                  background-color: #ffffff;
                }
                .performance-scorecard .table .spacer {
                  // width: 50px !important;
                  background-color: #cfd6de !important;
                  border-left: 1px solid #cfd6de !important;
                  border-bottom: 1px solid #cfd6de !important;
                  border-top: 1px solid #cfd6de !important;
                }
                .performance-scorecard .table .bdr-rt {
                  border-right: 1px solid #abd0ec;
                }
                .performance-scorecard .table thead {
                  position: sticky;
                  top: 0;
                  z-index: 99;
                  font-family: "FaktPro-Bold";
                  font-size: 0.8rem;
                }
                .performance-scorecard .table thead th {
                  padding: 5px 10px;
                  margin: 0px;
                }
                .performance-scorecard .table thead .bdr-btm {
                  border-bottom: #005fa3;
                }
                .performance-scorecard .table thead tr:nth-child(1) {
                  background-color: #fff !important;
                  color: #000;
                  text-transform: uppercase;
                  border-bottom: #cfd6de;
                }
                .performance-scorecard .table thead tr:nth-child(2) {
                  background-color: #337ab7 !important;
                  color: #fff;
                  text-transform: uppercase;
                  border-bottom: #cfd6de;
                  box-shadow: inset 0 1px 0 0 #cfd6de;
                }
                .performance-scorecard .table thead tr:nth-child(3) {
                  background-color: #337ab7 !important;
                  color: #fff;
                  text-transform: uppercase;
                  border-bottom: #cfd6de;
                  box-shadow: inset 0 1px 0 0 #cfd6de;
                }
                .performance-scorecard .table tbody {
                  font-family: "FaktPro-Normal";
                  font-size: 0.9rem;
                }
                .performance-scorecard .table tbody  td {
                  padding: 2px 10px;
                  margin: 0px;
                  border: 1px solid #cfd6de;
                }
                .performance-scorecard .table tbody tr {
                  border-bottom: 1px solid #37a6f8;
                  border-left: 1px solid #37a6f8;
                }
                .performance-scorecard .table tbody td:first-child {
                  text-align: start;
                  box-shadow: inset -1px 0 0 0 #cfd6de;
                }
                .performance-scorecard .table tbody .text-bold {
                  font-family: "FaktPro-Bold";
                }
                .performance-scorecard .table tbody .darkred-bg {
                  background-color: #282828 !important;
                  color: #fff;
                }
                .performance-scorecard .table tbody .lightblue-bg {
                  background-color: #646e7a !important;
                  color: #fff;
                }
                .performance-scorecard .table tbody  .gold-bg {
                  background-color: #ffa51a;
                  color: #fff;
                }
                </style>
                </head>
                <body id='content'>
                    ${printContents}
                </body>
            </html>
        `);

    doc.close();

    const div = doc.getElementById('content');
    const options = {
      logging: false,
      allowTaint: false,
      useCORS: true,
    };
    if (!div) {
      return;
    }
    html2canvas(div, options)
      .then((canvas) => {
        let imgWidth = 280;
        let pageHeight = 204;
        let imgHeight = (canvas.height * imgWidth) / canvas.width;
        let heightLeft = imgHeight;
        const contentDataURL = canvas.toDataURL('image/jpeg', 0.5);;
        let pdfData = new jsPDF('l', 'mm', 'a4', true);
        let position = 5;
        pdfData.addImage(
          contentDataURL,
          'JPEG',
          5,
          position,
          imgWidth,
          imgHeight,
          undefined,
          'FAST'
        );
        heightLeft -= pageHeight;
        while (heightLeft >= 0) {
          position = heightLeft - imgHeight;
          pdfData.addPage();
          pdfData.addImage(
            contentDataURL,
            'JPEG',
            5,
            position,
            imgWidth,
            imgHeight,
            undefined,
            'FAST'
          );
          heightLeft -= pageHeight;
        }

        const pdfBlob = pdfData.output('blob');
        const pdfFile = this.blobToFile(pdfBlob, 'LoanerInventory.pdf');
        const formData = new FormData();
        formData.append('to_email', Email);
        formData.append('subject', 'LoanerInventory');
        formData.append('file', pdfFile);
        formData.append('notes', notes);
        formData.append('from', from);
        this.Api.postmethod(this.comm.routeEndpoint + 'mail', formData).subscribe(
          (res: any) => {
            console.log('Response:', res);
            if (res.status === 200) {
              // alert(res.response);
              this.toast.success(res.response)
            } else {
              alert('Invalid Details');
            }
          },
          (error) => {
            console.error('Error:', error);
          }
        );
      })
      .catch((error) => {
        console.error('html2canvas error:', error);
      })
      .finally(() => {
        this.spinner.hide();
        // popupWin.close();
      });
  }
  public blobToFile = (theBlob: Blob, fileName: string): File => {
    return new File([theBlob], fileName, {
      lastModified: new Date().getTime(),
      type: theBlob.type,
    });
  };
}
