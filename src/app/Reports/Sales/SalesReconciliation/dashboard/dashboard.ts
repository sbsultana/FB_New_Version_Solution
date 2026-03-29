import { ChangeDetectorRef, Component, HostListener, Inject, PLATFORM_ID } from '@angular/core';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { BsDatepickerConfig, BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { CurrencyPipe, DatePipe, DecimalPipe } from '@angular/common';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
import { NgxSpinnerModule, NgxSpinnerService } from 'ngx-spinner';
import { HttpClient } from '@angular/common/http';
import { NgbModal } from '@ng-bootstrap/ng-bootstrap';
import { common } from '../../../../common';
import { FormBuilder } from '@angular/forms';
import { Router } from 'express';
import { Api } from '../../../../Core/Providers/Api/api';
import { ToastContainer } from '../../../../Layout/toast-container/toast-container';
import { Stores } from '../../../../CommonFilters/stores/stores';
import FileSaver from 'file-saver';
import { Workbook } from 'exceljs';
import { Subscription } from 'rxjs';
const EXCEL_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
const EXCEL_EXTENSION = '.xlsx';

@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [SharedModule, NgxSpinnerModule, BsDatepickerModule, Stores],
  providers: [CurrencyPipe, DatePipe, DecimalPipe],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
})
export class Dashboard {


  dynamicTitle = 'Sales Reconciliation';
  selectedMonth: any = new Date();
  bsConfig!: Partial<BsDatepickerConfig>;
  maxDate: Date = new Date();
  bsValue: Date = new Date();

  storeIds: any = [4]
  stores: any = []
  groupsArray: any = [];
  storename: any = ''
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  groupId: any = 0;
  groups: any = 1;
  storesFilterData: any = {
    'groupsArray': this.groupsArray, 'groupId': this.groupId, 'storesArray': this.stores, 'storeids': this.storeIds, 'type': 'S', 'others': 'N',
    'groupName': this.groupName, 'storename': this.storename, storecount: null, 'storedisplayname': this.storedisplayname
  };
  TotalReport: any = 'T';
  dateType: any = 'MTD';


  constructor(public shared: Sharedservice, private comm: common, private apiSrvc: Api, private spinner: NgxSpinnerService, private datepipe: DatePipe,private toast: ToastService,) {
    this.shared.setTitle('Sales Reconciliation');
    if (typeof window !== 'undefined') {
      if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
        // this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.ustores.split(',')[0]
        this.groupId = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Preferences
        this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids.split(',')[0]
      }
      if (this.shared.common.groupsandstores.length > 0) {
        this.groupsArray = this.shared.common.groupsandstores.filter((val: any) => val.sg_id != this.shared.common.reconID);
        this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
        this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_Name : this.groupName = ''
        this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
        this.getStoresandGroupsValues()
      }
      this.shared.setTitle(this.shared.common.titleName + '-Sales Reconciliation');
      const data = { title: this.dynamicTitle, stores: '2' };
      this.shared.api.SetHeaderData({ obj: data });
      this.bsConfig = {
        dateInputFormat: 'MMMM/YYYY',
        minMode: 'month',
        maxDate: this.maxDate
      };
      this.getSalesReconciliationData(this.selectedMonth);
    }
  }

  ngOnInit(): void { }

  StoresData(data: any) {
    this.storeIds = data.storeids;
    this.groupId = data.groupId;
    this.storename = data.storename;
    this.groupName = data.groupName;
    this.storecount = data.storecount;
    this.storedisplayname = data.storedisplayname;
  }
  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    }
    return false;
  }

  SPRstate: any;
  excel!: Subscription;
  ngAfterViewInit(): void {
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.shared.common.pageName == 'Sales Reconciliation') {
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
        if (res.obj.title == 'Sales Reconciliation') {
          if (res.obj.state == true) {
            this.exportAsXLSX();
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
      'type': 'S', 'others': 'N'
    };
  }



  onYearSelected(event: Date) {
    if (event) {
      const day = '01';
      const monthNames = [
        'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
        'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'
      ];
      const monthName = monthNames[event.getMonth()];
      const year = event.getFullYear();
      this.selectedMonth = `${day}-${monthName}-${year}`;
    }
  }

  SRData: any;
  SrdataXlsx: any = [];
  NoData: boolean = false;
  getSalesReconciliationData(Month: any): void {
    this.SRData = [1];
    let SubStoreArr: any = [];
    this.spinner.show();
    this.storename = this.stores
      .filter((val: any) => this.storeIds.includes(val.ID))
      .map((val: any) => val.storename)
      .join(', ');
    let Obj = {
      AS_ID: this.storeIds.toString(),
      DATE: this.shared.datePipe.transform(Month, 'dd-MMM-yyyy'),
      UserID: 0,
      // "AS_ID": "8",
      // "DATE": "30-Jan-2023",
      // "UserID": "1"
    };
    this.apiSrvc.postmethod(this.comm.routeEndpoint + 'GetSalesReconciliation', Obj).subscribe(
      (res) => {
        const currentTitle = document.title;
        if (res.status == 200) {
          this.SRData = res.response.reduce((r: any, { Lable1 }: any) => {
            if (!r.some((o: any) => o.Lable1 == Lable1)) {
              r.push({
                Lable1,
                sub: res.response.filter((v: any) => v.Lable1 == Lable1),
              });
            }
            return r;
          }, []);
          this.SRData.forEach((e: any, i: any) => {
            SubStoreArr.push(
              e.sub.reduce((r: any, { Lable1_subcategory }: any) => {
                if (
                  !r.some(
                    (o: any) => o.Lable1_subcategory == Lable1_subcategory
                  )
                ) {
                  r.push({
                    Lable1_subcategory,
                    subcategory: e.sub.filter(
                      (v: any) => v.Lable1_subcategory == Lable1_subcategory
                    ),
                  });
                }
                return r;
              }, [])
            );
            e.sub = SubStoreArr[i];
          });
          console.log('SR Data', this.SRData);
          this.spinner.hide();
          this.SrdataXlsx = [...res.response];
          if (this.SrdataXlsx.length > 0) {
            for (var i = 0; i < this.SrdataXlsx.length; i++) {
              if (
                this.SrdataXlsx[i]['Lable2'].toUpperCase() ==
                'Deal Adjustments'.toUpperCase()
              ) {
                this.SrdataXlsx.splice(i, 0, {
                  Lable2: 'Accounting Adjustments',
                });
                i++;
              }
            }
          }
        } else {
                
          this.toast.show('Invalid Details.', 'danger', 'Error');

          this.spinner.hide();
        }
      },
      (error) => {
        console.log(error);
      }
    );
  }

  //////////////////////  REPORT CODE /////////////////////////////////////////////
  activePopover: number = -1;
  custom: boolean = false;

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

  applyFilterAndYear() {
    this.activePopover = -1;

    if (this.storeIds && this.storeIds.length > 0) {
      const data = {
        Reference: 'Sales Reconciliation',
        storeValues: this.storeIds.toString(),
        groups: this.groupId.toString(),
      };
      this.shared.api.SetReports({
        obj: data,
      });
      console.log(data, this.selectedMonth);
      this.getSalesReconciliationData(this.selectedMonth);
    } else {
    
      this.toast.show('Please Select Atleast One Store', 'warning', 'Warning');
    }
  }

  formatCurrency(val: any) {
    if (val === null || val === undefined || val === 0) {
      return '-';
    }
    return '$' + Number(val).toLocaleString();
  }

  formatNumber(val: any) {
    if (val === null || val === undefined || val === 0) {
      return '-';
    }
    return Number(val).toLocaleString();
  }

  exportAsXLSX(): void {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet(
      'Sales Reconciliation'
    );
    worksheet.views = [
      {
        state: 'frozen',
        ySplit: 11, // Number of rows to freeze (2 means the first two rows are frozen)
        topLeftCell: 'A12', // Specify the cell to start freezing from (in this case, the third row)
        showGridLines: false,
      },
    ];
    worksheet.addRow('');
    const titleRow = worksheet.addRow(['Sales Reconciliation']);
    titleRow.eachCell((cell, number) => {
      cell.alignment = { indent: 1, vertical: 'middle', horizontal: 'left' };
    });
    titleRow.font = { name: 'Arial', family: 4, size: 12, bold: true };
    titleRow.worksheet.mergeCells('A2', 'D2');

    worksheet.addRow('');
    const DateToday = this.datepipe.transform(
      new Date(),
      'MM/dd/yyyy h:mm:ss a'
    );
    const dateMonth = this.datepipe.transform(
      new Date(this.selectedMonth),
      'MMMM yyyy'
    );

    worksheet.addRow([DateToday]).font = { name: 'Arial', family: 4, size: 9 };
    const SalesReconciliation = this.SrdataXlsx.map((_arrayElement: any) =>
      Object.assign({}, _arrayElement)
    );

    const Header = [
      dateMonth,
      'Units',
      'Front Gross',
      'Back Gross',
      'Total Gross',
      'Units',
      'Front Gross',
      'Back Gross',
      'Total Gross',
      'Units',
      'Front Gross',
      'Back Gross',
      'Total Gross',
    ];
    // for (let i = 0; i < this.lastFourMonthsArray.length; i++) {
    //  const formattedDate = this.datepipe.transform(this.lastFourMonthsArray[i], 'MMMM yyyy');
    //   Header.push(formattedDate!);
    // }
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
    datemonth.value = dateMonth;
    datemonth.font = { name: 'Arial', family: 4, size: 9 };
    datemonth.alignment = { vertical: 'middle', horizontal: 'left' };
    DateMonth.getCell(1).font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
    };
    const Stores = worksheet.addRow(['Stores :']);
    const stores = worksheet.getCell('B8');
    stores.value = this.storename;
    stores.font = { name: 'Arial', family: 4, size: 9 };
    stores.alignment = { vertical: 'middle', horizontal: 'left' };
    Stores.getCell(1).font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
    };
    worksheet.addRow('');

    const row = worksheet.getRow(10);
    row.height = 20;
    let Customer = worksheet.getCell('A10');
    Customer.value = this.storename;
    Customer.alignment = { vertical: 'middle', horizontal: 'center' };
    Customer.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
      color: { argb: 'FFFFFF' },
    };
    Customer.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '2a91f0' },
      bgColor: { argb: 'FF0000FF' },
    };
    Customer.border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' },
    };
    worksheet.mergeCells('B10', 'E10');

    let Tran = worksheet.getCell('B10');
    Tran.value = 'Transaction';
    Tran.alignment = { vertical: 'middle', horizontal: 'center' };
    Tran.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
      color: { argb: 'FFFFFF' },
    };
    Tran.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '2a91f0' },
      bgColor: { argb: 'FF0000FF' },
    };
    Tran.border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' },
    };
    worksheet.mergeCells('F10', 'I10');
    let Acct = worksheet.getCell('F10');
    Acct.value = 'Accounting';
    Acct.alignment = { vertical: 'middle', horizontal: 'center' };
    Acct.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
      color: { argb: 'FFFFFF' },
    };
    Acct.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '2a91f0' },
      bgColor: { argb: 'FF0000FF' },
    };
    Acct.border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' },
    };
    worksheet.mergeCells('J10', 'M10');
    let Diff = worksheet.getCell('J10');
    Diff.value = 'Difference';
    Diff.alignment = { vertical: 'middle', horizontal: 'center' };
    Diff.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
      color: { argb: 'FFFFFF' },
    };
    Diff.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '2a91f0' },
      bgColor: { argb: 'FF0000FF' },
    };
    Diff.border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' },
    };
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
        fgColor: { argb: '788494' },
        bgColor: { argb: 'FF0000FF' },
      };
      cell.border = { right: { style: 'thin' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });
    SalesReconciliation.forEach((d: any) => {
      // Update label values if necessary
      switch (d.Lable2) {
        case 'NewTotal':
          d.Lable2 = 'New';
          break;
        case 'UsedTotal':
          d.Lable2 = 'Used';
          break;
        case 'OtherTotal':
          d.Lable2 = 'Other';
          break;
        case 'GrandTotal':
          d.Lable2 = 'Total';
          break;
      }

      const Obj = [
        d.Lable2 || '-',
        d.Sale_Units || '-',
        d.Sale_Frontgross || '-',
        d.Sale_Backgross || '-',
        d.Sale_Totalgross || '-',
        d.Acct_Units || '-',
        d.Acct_Frontgross || '-',
        d.Acct_Totalgross || '-',
        d.Acct_Backgross || '-',
        d.Diff_Units || '-',
        d.Diff_Frontgross || '-',
        d.Diff_Backgross || '-',
        d.Diff_Totalgross || '-',
      ];
      const row = worksheet.addRow(Obj);
      row.font = { name: 'Arial', family: 4, size: 8 };

      row.eachCell((cell, number) => {
        cell.border = { right: { style: 'thin' } };
        if (number === 1) {
          cell.alignment = {
            indent: 2,
            vertical: 'middle',
            horizontal: 'left',
          };
        } else {
          cell.alignment = {
            vertical: 'middle',
            horizontal: 'right',
            indent: 1,
          };
        }
      });
      if (['New', 'Used', 'Other', 'Total'].includes(d.Lable2)) {
        row.getCell(1).alignment = {
          indent: 1,
          vertical: 'middle',
          horizontal: 'left',
        };
        row.font = { name: 'Arial', family: 4, size: 9, bold: true };
      } else if (d.Lable2 === 'Accounting Adjustments') {
        row.getCell(1).alignment = {
          indent: 1,
          vertical: 'middle',
          horizontal: 'left',
        };
        row.font = {
          name: 'Arial',
          family: 4,
          size: 9,
          bold: true,
          color: { argb: '2a91f0' },
        };
      } else {
        row.getCell(1).alignment = {
          indent: 2,
          vertical: 'middle',
          horizontal: 'left',
        };
      }
      row.eachCell((cell: any, number: any) => {
        if ([2, 6, 10].includes(number)) {
          cell.numFmt = '#,##0';
        } else {
          cell.numFmt = '$#,##0.00';
        }
      });
      if (row.number % 2) {
        row.eachCell((cell, number) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'e5e5e5' },
            bgColor: { argb: 'FF0000FF' },
          };
        });
      }
    });

    worksheet.eachRow((row, rowIndex) => {
      row.eachCell((cell, colIndex) => {
        if (rowIndex > 1 && rowIndex < 9) {
          if (colIndex === 1) {
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
    // worksheet.getColumn(1).alignment = {
    //   indent: 1,
    //   vertical: 'middle',
    //   horizontal: 'left',
    // };
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
    worksheet.addRow([]);
    workbook.xlsx.writeBuffer().then((data: any) => {
      const blob = new Blob([data], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      FileSaver.saveAs(
        blob,
        'Sales Reconciliation' + EXCEL_EXTENSION
      );
    });
  }
}