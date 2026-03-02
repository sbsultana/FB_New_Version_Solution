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
  // imports: [SharedModule, BsDatepickerModule, Stores, DateRangePicker],
  imports:[],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
})
export class Dashboard {
  // SalesTaxData: any = [];
  // FromDate: any = '';
  // ToDate: any = '';
  // minDate!: Date;
  // maxDate!: Date;
  // DateType: any = 'MTD';
  // displaytime: any = '';
  // NoData: any = '';

  // //  storeIds: any = []
  // stores: any = []
  // groupsArray: any = [];
  // storename: any = ''
  // storecount: any = null;
  // storedisplayname: any = '';
  // groupName: any = '';
  // groupId: any = 0;

  // storesFilterData: any = {
  //   'groupsArray': this.groupsArray, 'groupId': this.groupId, 'storesArray': this.stores, 'storeids': '1', 'type': 'M', 'others': 'N',
  //   'groupName': this.groupName, 'storename': this.storename, storecount: null, 'storedisplayname': this.storedisplayname
  // };

  // Dates: any = {
  //   'FromDate': this.FromDate, 'ToDate': this.ToDate, "MaxDate": this.maxDate, 'MinDate': this.minDate, 'DateType': this.DateType, 'DisplayTime': this.displaytime,
  //   Types: [
  //     { 'code': 'MTD', 'name': 'MTD' },
  //     { 'code': 'QTD', 'name': 'QTD' },
  //     { 'code': 'YTD', 'name': 'YTD' },
  //     { 'code': 'PYTD', 'name': 'PYTD' },
  //     { 'code': 'LY', 'name': 'Last Year' },
  //     { 'code': 'LM', 'name': 'Last Month' },
  //     { 'code': 'PM', 'name': 'Same Month PY' },
  //   ]
  // }
  

  // // filters
  // TotalReport: any = 'T';
  // storeIds!: any;
  // dateType: any = 'MTD';
  // groups: any = 1;
  // storeorgrp: any = 'G';
  // saleType: any = 'Retail,Lease,Wholesale,Misc,Fleet,Demo,Special Order,Rental,Dealer Trade';
  // retailorlease: any = this.saleType.split(',');
  // dealStatus: any = ['Booked', 'Finalized', 'Delivered'];

  // CompleteComponentState: boolean = true;


  // constructor(public shared: Sharedservice, public setdates: Setdates, private comm: common, private cp: CurrencyPipe,private toast: ToastService) {
  //   this.shared.setTitle('Sales Tax');
  //   this.initializeDates(this.DateType)


  //   if (typeof window !== 'undefined') {
  //     if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
  //       // this.storeIds = '1,2'
  //       this.groupId = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Preferences
  //       this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids.split(',')
  //     }
  //     if (this.shared.common.groupsandstores.length > 0) {
  //       this.groupsArray = this.shared.common.groupsandstores.filter((val: any) => val.sg_id != this.shared.common.reconID);
  //       this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
  //       this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_Name : this.groupName = ''
  //       this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
  //       this.getStoresandGroupsValues()
  //     }


  //     this.shared.setTitle(this.shared.common.titleName + '-Sales Tax');
  //     const data = {
  //       title: 'Sales Tax',
  //       stores: this.storeIds,
  //     };
  //     this.shared.api.SetHeaderData({
  //       obj: data,
  //     });

  //     this.GetData();
  //     this.setDates(this.DateType)
  //   }
  // }

  // ngOnInit(): void { }

  // getTotal(frontgross: any[], colname: string) {
  //   return frontgross.reduce((total, x) => {
  //     const raw = x?.[colname];
  //     // Skip null/undefined/empty strings
  //     if (raw === null || raw === undefined || raw === '') return total;
  //     // Parse as number (handles strings like "123" or "123.45")
  //     const n = Number(raw);
  //     // Skip NaN
  //     if (Number.isNaN(n)) return total;
  //     // Add the value (use Math.trunc if you only want integer part)
  //     return total + n;
  //   }, 0);
  // }

  // StoresData(data: any) {
  //   this.storeIds = data.storeids;
  //   this.groupId = data.groupId;
  //   this.storename = data.storename;
  //   this.groupName = data.groupName;
  //   this.storecount = data.storecount;
  //   this.storedisplayname = data.storedisplayname;
  // }

  // initializeDates(type: any) {
  //   let dates: any = this.setdates.setDates(type)
  //   this.FromDate = dates[0];
  //   this.ToDate = dates[1];
  //   localStorage.setItem('time', type);
  // }

  // GetData() {

  //   this.SalesTaxData = [];
  //   this.shared.spinner.show();
  //   const obj = {
  //     "AS_ID": this.storeIds.toString(),
  //     "FromDate": this.shared.datePipe.transform(this.FromDate, 'MM-dd-yyyy'),
  //     "ToDate": this.shared.datePipe.transform(this.ToDate, 'MM-dd-yyyy'),
  //   };
  //   const curl = this.shared.getEnviUrl() + 'GetSalesTax';
  //   this.shared.api.postmethod(this.shared.common.routeEndpoint + 'GetSalesTax', obj).subscribe((res) => {
  //     if (res.status == 200) {
  //       if (res.response != undefined) {
  //         if (res.response.length > 0) {
  //           this.SalesTaxData = res.response;
  //           this.shared.spinner.hide();
  //           this.NoData = '';
  //           console.log(this.SalesTaxData, 'Sales Tax Data.........');

  //         } else {
  //           this.shared.spinner.hide();
  //           this.NoData = 'No Data Found!!';
  //         }
  //       } else {
  //         this.shared.spinner.hide();
  //         this.NoData = 'No Data Found!!';
  //       }
  //     } else {
  //       this.shared.spinner.hide();
  //       this.NoData = 'No Data Found!!';
  //     }
  //   },
  //     (error) => {
  //       //  this.shared.toaster.error('502 Bad Gate Way Error', '');
  //       this.shared.spinner.hide();
  //       this.NoData = 'No Data Found!!';
  //     }
  //   );
  // }

  // public inTheGreen(value: number): boolean {
  //   if (value >= 0) {
  //     return true;
  //   }
  //   return false;
  // }

  // isDesc: boolean = false;
  // column: string = 'CategoryName';
  // sort(property: any) {
  //   this.isDesc = !this.isDesc; //change the direction
  //   this.column = property;
  //   let direction = this.isDesc ? 1 : -1;

  //   this.SalesTaxData.sort((a: any, b: any) => {
  //     const valA = a[property] ?? ''; // replace null/undefined with empty string
  //     const valB = b[property] ?? '';

  //     if (valA < valB) {
  //       return -1 * direction;
  //     } else if (valA > valB) {
  //       return 1 * direction;
  //     } else {
  //       return 0;
  //     }
  //   });

  // }
  // SPRstate: any;
  // ngAfterViewInit(): void {
  //   this.shared.api.getStores().subscribe((res: any) => {
  //     if (this.comm.pageName == 'Sales Tax') {
  //       if (res.obj.storesData != undefined) {
  //         this.groupsArray = res.obj.storesData;
  //         this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
  //         this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_name : this.groupName = ''
  //         this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
  //         this.getStoresandGroupsValues()
  //       }
  //     }
  //   })

  //   this.shared.api.getExportToExcelAllReports().subscribe((res) => {
  //     this.SPRstate = res.obj.state;
  //     if (res.obj.title == 'Sales Tax') {
  //       if (res.obj.state == true) {
  //         this.exportToExcel();
  //       }
  //     }
  //   });

  //   this.shared.api.getExportToPrintAllReports().subscribe((res) => {
  //     if (res.obj.title == 'Sales Tax') {
  //       if (res.obj.statePrint == true) {
  //         // this.GetPrintData();
  //       }
  //     }
  //   });

  //   this.shared.api.getExportToPDFAllReports().subscribe((res) => {
  //     if (res.obj.title == 'Sales Tax') {
  //       if (res.obj.statePDF == true) {
  //         // this.generatePDF();
  //       }
  //     }
  //   });
  //   this.shared.api.getExportToEmailPDFAllReports().subscribe((res) => {
  //     if (res.obj.title == 'Sales Tax') {
  //       if (res.obj.stateEmailPdf == true) {
  //         // this.sendEmailData(res.obj.Email, res.obj.notes, res.obj.from);
  //       }
  //     }
  //   });
  // }


  // getStoresandGroupsValues() {
  //   this.storesFilterData.groupsArray = this.groupsArray;
  //   this.storesFilterData.groupId = this.groupId;
  //   this.storesFilterData.storesArray = this.stores;
  //   this.storesFilterData.storeids = this.storeIds;
  //   this.storesFilterData.groupName = this.groupName;
  //   this.storesFilterData.storename = this.storename;
  //   this.storesFilterData.storecount = this.storecount;
  //   this.storesFilterData.storedisplayname = this.storedisplayname;
  //   this.storesFilterData = {
  //     groupsArray: this.groupsArray,
  //     groupId: this.groupId,
  //     storesArray: this.stores,
  //     storeids: this.storeIds,
  //     groupName: this.groupName,
  //     storename: this.storename,
  //     storecount: this.storecount,
  //     storedisplayname: this.storedisplayname,
  //     'type': 'M', 'others': 'N'
  //   };
  // }



  // openCustomPicker(datepicker: any) {
  //   this.DateType = 'C';
  //   localStorage.setItem('time', this.DateType);
  //   datepicker.toggle();
  // }


  // dateRangeCreated($event: any) {
  //   if ($event) {
  //     this.FromDate = this.shared.datePipe.transform($event[0], 'MM-dd-yyyy');
  //     this.ToDate = this.shared.datePipe.transform($event[1], 'MM-dd-yyyy');
  //     this.bsRangeValue = [this.FromDate, this.ToDate];
  //     if (this.DateType === 'C') this.custom = true;
  //   }
  // }

  // updatedDates(data: any) {
  //   // console.log(data);
  //   this.FromDate = data.FromDate;
  //   this.ToDate = data.ToDate;
  //   this.DateType = data.DateType;
  //   this.displaytime = data.DisplayTime
  // }




  // // comments code

  // Scrollpercent: any = 0;
  // scrollCurrentposition: any = 0;
  // @ViewChild('scrollcent') scrollcent!: ElementRef;
  // updateVerticalScroll(event: any): void {
  //   this.scrollCurrentposition = event.target.scrollTop;
  //   const scrollDemo = document.querySelector('#scrollcent') as HTMLElement;
  //   this.Scrollpercent = Math.round(
  //     (event.target.scrollTop / (event.target.scrollHeight - scrollDemo.clientHeight)) * 100
  //   );
  // }






  // //////////////////////  REPORT CODE /////////////////////////////////////////////
  // activePopover: number = -1;
  // custom: boolean = false;
  // bsRangeValue!: Date[];
  // @HostListener('document:click', ['$event'])
  // onDocumentClick(event: MouseEvent) {
  //   const clickedInside = (event.target as HTMLElement).closest('.dropdown-toggle, .reportstores-card, .timeframe');
  //   if (!clickedInside) {
  //     this.activePopover = -1;
  //   }
  // }

  // togglePopover(popoverIndex: number) {
  //   this.activePopover = this.activePopover === popoverIndex ? -1 : popoverIndex;
  // }


  // setDates(type: any) {
  //   this.displaytime = '(' + this.Dates.Types.filter((val: any) => val.code == type)[0].name + ')';
  //   this.maxDate = new Date();
  //   this.minDate = new Date();
  //   this.minDate.setFullYear(this.maxDate.getFullYear() - 3);
  //   this.maxDate.setDate(this.maxDate.getDate());
  //   this.Dates.FromDate = this.FromDate;
  //   this.Dates.ToDate = this.ToDate;
  //   this.Dates.MinDate = this.minDate;
  //   this.Dates.MaxDate = this.maxDate;
  //   this.Dates.DateType = this.DateType;
  //   this.Dates.DisplayTime = this.displaytime;
  // }
  // close() {
  //   this.shared.ngbmodal.dismissAll();
  // }

  // // Final Apply
  // viewreport() {
  //   this.activePopover = -1;

  //   if (this.storeIds && this.storeIds.length > 0) {
  //     const data = {
  //       Reference: 'Sales Tax',
  //       FromDate: this.FromDate,
  //       ToDate: this.ToDate,
  //       storeValues: this.storeIds.toString(),
  //       dateType: this.DateType,
  //       groups: this.groupId.toString(),
  //       saleType: this.retailorlease.toString(),
  //       dealStatus: this.dealStatus,
  //     };
  //     this.shared.api.SetReports({
  //       obj: data,
  //     });
  //     this.close();
  //     this.GetData();
  //   } else {
  
  //     this.toast.show('Please select atleast one store', 'warning', 'Warning');
  //   }




  // }
  // exportToExcel(): void {
  //   const workbook = this.shared.getWorkbook();
  //   const worksheet = workbook.addWorksheet('Sales Tax');
  //   const title = worksheet.addRow(['Sales Tax']);
  //   title.font = { size: 14, bold: true, name: 'Arial' };
  //   title.alignment = { vertical: 'middle', horizontal: 'left' };
  //   worksheet.mergeCells('A1:L1');
  //   worksheet.addRow([]);
  //   const formattedFromDate = this.shared.datePipe.transform(this.FromDate, 'MM/dd/yyyy');
  //   const formattedToDate = this.shared.datePipe.transform(this.ToDate, 'MM/dd/yyyy');
  //   let storeValue = 'All Stores';
  //   // if (
  //   //   this.storeIds &&
  //   //   this.storeIds.length > 0 &&
  //   //   this.storeIds.length !== this.stores.length
  //   // ) {
  //   storeValue = this.stores
  //     .filter((s: any) => this.storeIds.includes(s.ID))
  //     .map((s: any) => s.storename)
  //     .join(', ');
  //   // }
  //   const filters = [
  //     { name: 'Store:', values: storeValue },
  //     { name: 'Time Frame:', values: `${formattedFromDate} to ${formattedToDate}` },
  //   ];
  //   let currentRow = worksheet.lastRow?.number ?? worksheet.rowCount;
  //   filters.forEach((filter) => {
  //     currentRow++;
  //     let value = Array.isArray(filter.values) ? filter.values.join(', ') : filter.values;

  //     const row = worksheet.addRow([filter.name, value]);
  //     row.getCell(1).font = { bold: true, name: 'Arial', size: 10 };
  //     // row.getCell(2).font = { name: 'Arial', size: 10, color: { argb: 'FF1F497D' } }; // blue color for values
  //     worksheet.mergeCells(`B${currentRow}:F${currentRow}`);
  //   });

  //   worksheet.addRow([]);

  //   const ReportTotals = ['', '', '', '', '', '', '', '', '', '', '', '', '', 'Report Totals', this.getTotal(this.SalesTaxData, 'Taxable CCR'), '',
  //     this.getTotal(this.SalesTaxData, 'SALE AMT'), this.getTotal(this.SalesTaxData, 'DOC FEES'), this.getTotal(this.SalesTaxData, 'TOTAL TAXABLE'),
  //   '', this.getTotal(this.SalesTaxData, 'TAX FROM 324'), this.getTotal(this.SalesTaxData, '7.25%'),
  //     this.getTotal(this.SalesTaxData, 'DISTRICT TAX'), this.getTotal(this.SalesTaxData, 'DIFF'), this.getTotal(this.SalesTaxData, 'DIFF 9.750%')];

  //   const totalRow = worksheet.addRow(ReportTotals);
  //   totalRow.font = { bold: true, color: { argb: '000000' } };
  //   totalRow.alignment = { vertical: 'middle', horizontal: 'right' };
  //   // headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2F5597' } };
  //   totalRow.height = 25;

  //   totalRow.eachCell((cell, colNumber) => {

  //     if (colNumber != 20) {
  //       cell.numFmt = '"$"#,##0.00'
  //     }
  //     // if (colNumber == 20) {
  //     //   cell.numFmt = '"$"#,##0.00'
  //     // }

  //     if (colNumber <= ReportTotals.length) { // Only style cells with header text
  //       cell.fill = {
  //         type: 'pattern',
  //         pattern: 'solid',
  //         fgColor: { argb: 'FFE8EEF7' } // Blue background
  //       };
  //     }
  //   });





  //   const firstHeader = ['Store Name', 'Status', 'Deal Type', 'Stock Type ', 'Contract Date', 'Final Accounting Date', 'Deal #', 'Stock #', 'Customer', 'Byer Street Address', 'Buyer City', 'Buyer State', 'Buyer Zip Code', 'Payment Type', 'Taxable CCR', 'Journal', 'Sale Amt', 'Doc Fees', 'Total Taxable', 'Rate', 'Tax From 324', '7.25%', 'District Tax', 'Diff', 'Diff 9.750%'];


  //   // worksheet.addRow(firstHeader);
  //   // worksheet.getRow(6).height = 25;
  //   // worksheet.getRow(6).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  //   // worksheet.getRow(6).alignment = { vertical: 'middle', horizontal: 'center' };
  //   // worksheet.getRow(6).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2F5597' } };

  //   const headerRow = worksheet.addRow(firstHeader);
  //   headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  //   headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
  //   // headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2F5597' } };
  //   headerRow.height = 25;

  //   headerRow.eachCell((cell, colNumber) => {
  //     if (colNumber <= firstHeader.length) { // Only style cells with header text
  //       cell.fill = {
  //         type: 'pattern',
  //         pattern: 'solid',
  //         fgColor: { argb: 'FF2F5597' } // Blue background
  //       };
  //     }
  //   });



  //   const bindingHeaders = [
  //     'Store', 'Status', 'Deal Type', 'Stock Type', 'Contract Date', 'Final accounting date', 'Deal Number', 'Stock#', 'Customer', 'Buyer Street Address',
  //     'Buyer City', 'Buyer State', 'Buyer Zip Code', 'Payment Type', 'Taxable CCR', 'Jrn.', 'SALE AMT', 'DOC FEES', 'TOTAL TAXABLE', 'RATE', 'TAX FROM 324', '7.25%', 'DISTRICT TAX', 'DIFF', 'DIFF 9.750%'
  //   ];

  //   const currencyFields = ['Taxable CCR', 'SALE AMT', 'DOC FEES', 'TOTAL TAXABLE', 'TAX FROM 324', '7.25%', 'DISTRICT TAX', 'DIFF', 'DIFF 9.750%'];

  //   this.SalesTaxData.forEach((info: any) => {
  //     const rowData = bindingHeaders.map((key) => {
  //       const val = info[key];
  //       if (key === 'Contract Date' || key === 'Final accounting date') return this.shared.datePipe.transform(val, 'MM/dd/yyyy');
  //       if (key == 'RATE') return (val === 0 || val == null || val === '') ? '-' : val + '%';
  //       return (val === 0 || val == null || val === '') ? '-' : val;
  //     });

  //     const dataRow = worksheet.addRow(rowData);

  //     bindingHeaders.forEach((key, index) => {
  //       const cell = dataRow.getCell(index + 1);

  //       if (currencyFields.includes(key) && typeof cell.value === 'number') {
  //         cell.numFmt = '"$"#,##0.00';
  //         cell.alignment = { horizontal: 'right', vertical: 'middle' };
  //       } else {
  //         cell.alignment = { horizontal: 'center', vertical: 'middle' };
  //       }
  //     });
  //   });


  //   worksheet.columns.forEach(col => col.width = 25);


  //   workbook.xlsx.writeBuffer().then(buffer => {
  //     this.shared.exportToExcel(workbook, 'Sales Tax');
  //   });
  // }
}
