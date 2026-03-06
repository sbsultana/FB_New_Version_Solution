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
import { NgbActiveModal, NgbModal, NgbModalRef } from '@ng-bootstrap/ng-bootstrap';
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
import { FormsModule, ReactiveFormsModule } from '@angular/forms';
import { Setdates } from '../../../../Core/Providers/SetDates/setdates';
import { DateRangePicker } from '../../../../CommonFilters/date-range-picker/date-range-picker';
import { Notes } from '../../../../Layout/notes/notes';
const EXCEL_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
const EXCEL_EXTENSION = '.xlsx';

@Component({
  selector: 'app-dashboard',
  imports: [CommonModule, SharedModule, Stores, FormsModule, ReactiveFormsModule,Notes],
  standalone: true,
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
})
export class Dashboard {
  //  storeIds: any = []
  stores: any = []
  groupsArray: any = [];
  storename: any = ''
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  groupId: any = 0;
  storeIds!: any;
  dateType: any = 'MTD';
  groups: any = 0;
  storesFilterData: any = {
    'groupsArray': this.groupsArray, 'groupId': this.groupId, 'storesArray': this.stores, 'storeids': '1', 'type': 'M', 'others': 'N',
    'groupName': this.groupName, 'storename': this.storename, storecount: null, 'storedisplayname': this.storedisplayname
  };

  constructor(public shared: Sharedservice, public setdates: Setdates, private comm: common, private toast: ToastService,
    public spinner: NgxSpinnerService, private Api: Api, private ngbmodal: NgbModal, private datepipe: DatePipe
  ) {
    this.shared.setTitle('Title Report');
    if (typeof window !== 'undefined') {
      if (localStorage.getItem('flag') == 'V') {
        this.storeIds = [];
        console.log(JSON.parse(localStorage.getItem('userInfo')!), JSON.parse(localStorage.getItem('userInfo')!).user_Info, 'Widget Stores............');
        this.groupId = JSON.parse(localStorage.getItem('userInfo')!).groupid
        JSON.parse(localStorage.getItem('userInfo')!).store.indexOf(',') > 0 ?
          this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).store.split(',') :
          this.storeIds.push(JSON.parse(localStorage.getItem('userInfo')!).store)
        localStorage.setItem('flag', 'M')
      } else {
        if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
          this.groupId = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Preferences
          this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids.split(',')
        }
      }
      if (this.shared.common.groupsandstores.length > 0) {
        this.groupsArray = this.shared.common.groupsandstores.filter((val: any) => val.sg_id != this.shared.common.reconID);
        this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
        this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_Name : this.groupName = ''
        this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
      }
      this.shared.setTitle(this.shared.common.titleName + '-Title Report');
      const data = {
        title: 'Title Report',
        stores: this.storeIds,
        groups: this.groups,
        count: 0,
      };
      this.shared.api.SetHeaderData({
        obj: data,
      });
    }
    this.GetTitle();
  }
  Scrollpercent: any = 0;
  scrollCurrentposition: any = 0;
  scrollpositionstoring: any = 0;
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
  StoresData(data: any) {
    console.log(data, 'Data');
    this.storeIds = data.storeids;
    this.groupId = data.groupId;
    this.storename = data.storename;
    this.groupName = data.groupName;
    this.storecount = data.storecount;
    this.storedisplayname = data.storedisplayname;
  }

  TitleData: any = [];
  TitleCompleteData: any = []
  SearchText: any = '';
  FromDate: any = '';
  ToDate: any = '';
  statustype: any = ['T', 'H', 'P', 'R', 'O', 'S', 'I']
  NoData: any = false;
  callLoadingState: any = 'FL';
  currentPage: number = 1;
  itemsPerPage: number = 100;
  maxPageButtonsToShow: number = 3;
  clickedPage: number | null = null;
  CurrentPageSetting: number = 1;
  GetTitle() {
    this.TitleData = [];
    this.TitleCompleteData = [];
    this.NoData = false;
    this.spinner.show();
    const obj = {
      'Store': this.storeIds.toString(),
      "StatusType": this.statustype.toString()
    };
    this.Api.postmethod(this.comm.routeEndpoint + 'GetTitleReport', obj).subscribe(
      (res) => {
        if (res.status == 200) {
          this.TitleData = [];
          if (res.response != undefined) {
            if (res.response.length > 0) {
              this.TitleCompleteData = res.response.map((v: any) => ({
                ...v, notesView: '+'
              }));
              this.TitleCompleteData.some(function (x: any) {
                if (x.Notes != undefined && x.Notes != null) {
                  x.Notes = JSON.parse(x.Notes);
                }
              });
              this.TitleData = this.TitleCompleteData;

              this.filterGrid(this.SearchText);
              console.log(this.TitleData);
              this.spinner.hide();
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

  isDesc: boolean = false;
  column: string = '';

  sort(property: any, data: any, state?: any) {
    if (state == undefined) {
      this.isDesc = !this.isDesc;
    }
    this.callLoadingState = 'FL'
    this.column = property;
    let direction = this.isDesc ? 1 : -1;
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

  get paginatedItems() {
    const startIndex = (this.currentPage - 1) * this.itemsPerPage;
    const endIndex = startIndex + this.itemsPerPage;
    return this.TitleData?.slice(startIndex, endIndex) || [];
  }
  getMaxPageNumber(): number {
    if (!this.TitleData || this.TitleData.length === 0) return 1;
    return Math.ceil(this.TitleData.length / this.itemsPerPage);
  }
  nextPage() {
    if (this.currentPage < this.getMaxPageNumber()) {
      this.currentPage++;
      this.clickedPage = null;
    }
  }
  prevPage() {
    if (this.currentPage > 1) {
      this.currentPage--;
      this.clickedPage = null;
    }
  }
  goToFirstPage() {
    this.currentPage = 1;
    this.clickedPage = null;
  }
  goToLastPage() {
    this.currentPage = this.getMaxPageNumber();
    this.clickedPage = null;
  }

  goToPage(page: number) {
    if (page >= 1 && page <= this.getMaxPageNumber()) {
      this.currentPage = page;
      this.clickedPage = page;
    }
  }
  onPageSizeChange() {
    this.currentPage = 1; 
  }
  getPageNumbers(): number[] {
    const maxPage = this.getMaxPageNumber();
    if (maxPage <= 0) return [];

    const current = this.currentPage;
    const maxButtons = this.maxPageButtonsToShow;

    let start = Math.max(current - Math.floor(maxButtons / 2), 1);
    let end = start + maxButtons - 1;

    if (end > maxPage) {
      end = maxPage;
      start = Math.max(end - maxButtons + 1, 1);
    }

    return Array.from({ length: end - start + 1 }, (_, i) => start + i);
  }

  getStartRecordIndex(): number {
    if (!this.TitleData || this.TitleData.length === 0) return 0;
    return (this.currentPage - 1) * this.itemsPerPage + 1;
  }

  getEndRecordIndex(): number {
    if (!this.TitleData || this.TitleData.length === 0) return 0;

    const endIndex = this.currentPage * this.itemsPerPage;
    return endIndex > this.TitleData.length
      ? this.TitleData.length
      : endIndex;
  }

  filterGrid(e: any) {
    console.log(e, '.............');

    this.SearchText = e
    if (this.SearchText.trim() !== '') {
      this.TitleData = this.TitleCompleteData.filter((item: any) =>
        (item.stockno && item.stockno.toLowerCase().includes(this.SearchText.toLowerCase())));
    } else {
      this.TitleData = [...this.TitleCompleteData];
    }
    this.TitleData.length > 0 ? this.NoData = false : this.NoData = true;
    this.callLoadingState == 'ANS' ? this.sort(this.column, this.TitleData, this.callLoadingState) : ''
    let position = this.scrollpositionstoring + 10
    setTimeout(() => {
      this.scrollcent.nativeElement.scrollTop = position
    }, 500);
  }

  toggleView(data: any) {
    if (data.notesView == '+') {
      data.notesView = '-'
    } else {
      data.notesView = '+'
    }
  }
  getTotal(columnname: any) {
    let total = 0
    this.TitleData.some(function (x: any) {
      total += x[columnname]
    })
    return total
  }
  closeNotes(e: any) {
    this.ngbmodal.dismissAll()

    // this.popup.close()
    if (e == 'S') {
      // alert(e)
      this.CurrentPageSetting = this.currentPage
      this.callLoadingState = 'ANS';
      this.notesViewState = true;
      this.GetTitle();
    }
  }
  notesViewState: boolean = true
  notesView() {
    this.notesViewState = !this.notesViewState
  }
  notesData: any = {}
  Notespopup: any
  addNotes(data: any, ref: any) {
    this.scrollpositionstoring = this.scrollCurrentposition;
    this.notesData = {
      store: data.inventorycompany,
      title1: data.stockno,
      title2: '',
      module: 'TR',
      apiRoute: 'AddGeneralNotes'
    }
    this.Notespopup = this.ngbmodal.open(ref, { size: 'xxl', backdrop: 'static' });
  }

  activePopover: number = -1;
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
  multipleorsingle(block: any, e: any) {
    const index = this.statustype.findIndex((i: any) => i == e);
    if (index >= 0) {
      this.statustype.splice(index, 1);
      if (this.statustype.length == 0) {
        this.toast.show('Please select atleast any one statustype', 'warning', 'Warning')
      }
    } else {
      this.statustype.push(e);
    }
  }
  viewreport() {
    this.activePopover = -1
    if (this.storeIds.length == 0) {
      this.toast.show('Please select atleast one store', 'warning', 'Warning');
    }
    else {
      this.GetTitle();
      this.currentPage = 1;
      this.GetTitle()
    }
  }

  LMstate: any;
  Pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;
  ngAfterViewInit() {

    this.shared.api.getStores().subscribe((res: any) => {
      if (this.comm.pageName == 'Title Report') {
        if (res.obj.storesData != undefined) {
          this.groupsArray = res.obj.storesData;
          this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
          this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_name : this.groupName = ''
          this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
          this.getStoresandGroupsValues()
        }
      }
    })
    this.email = this.Api.getExportToEmailPDFAllReports().subscribe((res) => {
      if (this.email != undefined) {
        if (res.obj.title == 'Title Report') {
          if (res.obj.stateEmailPdf == true) {
            this.sendEmailData(res.obj.Email, res.obj.notes, res.obj.from);
          }
        }
      }
    });
    this.print = this.Api.getExportToPrintAllReports().subscribe((res) => {
      if (this.print != undefined) {
        if (res.obj.title == 'Title Report') {
          if (res.obj.statePrint == true) {
            this.GetPrintData();
          }
        }
      }
    });
    this.Pdf = this.Api.getExportToPDFAllReports().subscribe((res) => {
      if (this.Pdf != undefined) {
        if (res.obj.title == 'Title Report') {
          if (res.obj.statePDF == true) {
            this.generatePDF();
          }
        }
      }
    });

    this.excel = this.Api.getExportToExcelAllReports().subscribe((res) => {
      if (this.excel != undefined) {
        this.LMstate = res.obj.state;
        if (res.obj.title == 'Title Report') {
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
  GetPrintData() {
    window.print();
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
  ExcelStoreNames: any = [];
  exportToExcel() {
    let storeNames: any = [];
    let store = this.storeIds;
    storeNames = this.comm.groupsandstores.filter((v: any) => v.sg_id == this.groups)[0].Stores.filter((item: any) => this.storeIds.includes(item.ID));
    if (store.length == this.comm.groupsandstores.filter((v: any) => v.sg_id == this.groups)[0].Stores.length) {
      this.ExcelStoreNames = 'All Stores'
    } else {
      this.ExcelStoreNames = storeNames.map(function (a: any) {
        return a.storename;
      });
    }


    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet('Title Report');
    worksheet.views = [
      {
        state: 'frozen',
        ySplit: 11, // Number of rows to freeze (2 means the first two rows are frozen)
        topLeftCell: 'A12', // Specify the cell to start freezing from (in this case, the third row)
        showGridLines: false,
      },
    ];
    worksheet.addRow('');
    const titleRow = worksheet.addRow(['Title Report']);
    titleRow.eachCell((cell, number) => {
      cell.alignment = { indent: 1, vertical: 'top', horizontal: 'left' };
    });
    titleRow.font = { name: 'Arial', family: 4, size: 12, bold: true };
    titleRow.worksheet.mergeCells('A2', 'D2');

    worksheet.addRow('');
    const DateToday = this.datepipe.transform(
      new Date(),
      'MM/dd/yyyy h:mm:ss a'
    );
    const DATE_EXTENSION = this.datepipe.transform(new Date(), 'MMddyyyy');
    worksheet.addRow([DateToday]).font = { name: 'Arial', family: 4, size: 9 };
    const Appointmentdata = this.TitleData.map((_arrayElement: any) =>
      Object.assign({}, _arrayElement)
    );
    const ReportFilter = worksheet.addRow(['Report Filters :']);
    ReportFilter.font = { name: 'Arial', family: 4, size: 10, bold: true };
    const Groups = worksheet.getCell('A6');
    Groups.value = 'Group :';
    Groups.font = { name: 'Arial', family: 4, size: 9, bold: true };
    const groups = worksheet.getCell('B6');
    groups.value =
      this.comm.groupsandstores.filter((val: any) => val.sg_id == this.groups.toString())[0].sg_name;
    groups.font = { name: 'Arial', family: 4, size: 9 };
    groups.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
    worksheet.mergeCells('B7', 'K9');
    const Stores = worksheet.getCell('A7');
    Stores.value = 'Stores :'
    const stores = worksheet.getCell('B7');
    stores.value = this.ExcelStoreNames == 0
      ? '-'
      : this.ExcelStoreNames == null
        ? '-'
        : this.ExcelStoreNames.toString().replaceAll(',', ', ');
    stores.font = { name: 'Arial', family: 4, size: 9 };
    stores.alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    Stores.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
    };
    const Timeframe = worksheet.addRow(['Timeframe :']);
    Timeframe.getCell(1).font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
    };
    const timeframe = worksheet.getCell('B10');
    timeframe.font = { name: 'Arial', family: 4, size: 9 };
    worksheet.addRow('');
    // const Note = worksheet.addRow(['Note :']);const note = worksheet.getCell("C4");note.value = this.StoreValues;note.font = {name: 'Arial',family: 4,size: 8,bold: false,};
    // Note.font = {name: 'Arial',family: 4,size: 8,bold: true,};
    let Headings = [
      'Stock #', 'Age',
      'Status',
      'Price 1',
      'Price 2', 'Store Name', 'Year', 'Make', 'Model', 'VIN', 'Deal Dept', 'Title Department', 'Title #', 'Balance'
    ];
    const headerRow = worksheet.addRow(Headings);
    headerRow.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: false,
      color: { argb: 'FFFFFF' },
    };
    headerRow.alignment = {
      indent: 1,
      vertical: 'middle',
      horizontal: 'center',
    };
    headerRow.eachCell((cell: any, number: any) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '788494' },
        bgColor: { argb: 'FF0000FF' },
      };
      cell.border = { right: { style: 'dotted' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });
    const SalesData = this.TitleData.map((_arrayElement: any) =>
      Object.assign({}, _arrayElement)
    );
    let count = 13

    for (const d of SalesData) {
      count++
      var data1obj = [
        d.stockno == '' ? '-' : d.stockno == null ? '-' : d.stockno,
        d.Age == '' ? '-' : d.Age == null ? '-' : d.Age,
        d.status == '' ? '-' : d.status == null ? '-' : d.status,
        d['price1'] == '' ? '-' : d['price1'] == null ? '-' : parseFloat(d['price1']),
        d['price2'] == '' ? '-' : d['price2'] == null ? '-' : parseFloat(d['price2']),
        d.StoreName == '' ? '-' : d.StoreName == null ? '-' : d.StoreName,
        d.Year == '' ? '-' : d.Year == null ? '-' : d.Year,
        d.Make == '' ? '-' : d.Make == null ? '-' : d.Make,
        d.Model == '' ? '-' : d.Model == null ? '-' : d.Model,
        d.VIN == '' ? '-' : d.VIN == null ? '-' : d.VIN,

        d['Deal Dept'] == '' ? '-' : d['Deal Dept'] == null ? '-' : d['Deal Dept'],
        d['Title Dept 2'] == '' ? '-' : d['Title Dept 2'] == null ? '-' : d['Title Dept 2'],
        d['Title Dept 1'] == '' ? '-' : d['Title Dept 1'] == null ? '-' : d['Title Dept 1'],

        d['balance'] == '' ? '-' : d['balance'] == null ? '-' : parseFloat(d['balance']),

      ]
      const Data1 = worksheet.addRow(data1obj);
      Data1.font = { name: 'Arial', family: 4, size: 9 };
      Data1.alignment = { vertical: 'middle', horizontal: 'center' };
      Data1.getCell(1).alignment = {
        indent: 1,
        vertical: 'middle',
        horizontal: 'left',
      };
      Data1.eachCell((cell: any, number: any) => {
        cell.border = { right: { style: 'dotted' } };
        // cell.numFmt = '$#,##0.00';

        if (number == 4 || number == 5 || number == 14) {
          cell.numFmt = '$#,##0.00';
        }
        if (number > 1 && data1obj[number] != undefined) {
          if (data1obj[number] < 0) {
            Data1.getCell(number + 1).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } };
          }
        }
      });
      if (d.NotesStatus == 'Y' && this.notesViewState == true) {
        for (const d1 of d.Notes) {
          worksheet.mergeCells(count, 1, count, 14);
          const Data2 = worksheet.getCell(count, 1);
          Data2.value = ' ' + ' ' + ' ' + ' ' + ' ' + ' ' + d1.GN_Text
          Data2.alignment = { indent: 2, vertical: 'middle', horizontal: 'left', };
          Data2.font = { name: 'Arial', family: 4, size: 9 };
          Data2.border = { right: { style: 'thin' }, left: { style: 'thin' }, top: { style: 'thin' }, bottom: { style: 'thin' } };
          Data2.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'b6d3ec' },
            bgColor: { argb: 'b4c7dc' },
          };
          count++
        }
      }
    }
    worksheet.getColumn(1).width = 15; worksheet.getColumn(1).alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
    worksheet.getColumn(2).width = 15;
    worksheet.getColumn(3).width = 20;
    worksheet.getColumn(4).width = 25;
    worksheet.getColumn(5).width = 25;
    worksheet.getColumn(6).width = 25;
    worksheet.getColumn(7).width = 25;

    worksheet.getColumn(8).width = 25;

    worksheet.getColumn(9).width = 25;
    worksheet.getColumn(10).width = 25;

    worksheet.getColumn(11).width = 25;
    worksheet.getColumn(12).width = 25;
    worksheet.addRow([]);
    workbook.xlsx.writeBuffer().then((data: any) => {
      const blob = new Blob([data], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      FileSaver.saveAs(
        blob,
        'Title Report_' + DATE_EXTENSION + EXCEL_EXTENSION
      );
    });
  }



  generatePDF() {
    this.spinner.show();
    const printContents = document.getElementById('TitleReport')!.innerHTML;
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
    <title>Title Report</title>
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
                 .negative {
                   color: red;
                 }
                 .HideItems{
                   display:none;
                  }            
   .backgrossbtn {
                   float: left;
                   font-size: 11px;
                   background-color: #2d91f1;
                   color: white;
                   padding: 4px;
                   border: 1px solid #2d91f1;
                   cursor: pointer;
                   border-radius: 3px;
               }
                 .Selectedrow {
                   background-color: #d3ecff !important;
                   color: #000 !important;
               }
                 .justify-content-between {
                   justify-content: space-between !important;
               }
               .d-flex {
                   display: flex !important;
               }
                 .bg-white {
                   background: #ffffff !important;
               }
                 .negative {
                   color: red;
                 }
                 .negative {
                   color: red;
                 }
                 .performance-scorecard .table>:not(:first-child) {
                   border-top: 0px solid #ffa51a
                 }
                 .performance-scorecard .table {
                   text-align: center;
                   text-transform: capitalize;
                   border: transparent;
                   width: 100%;
                 }
                 .performance-scorecard .table th,
                 .performance-scorecard .table td{
                   white-space: nowrap;
                   vertical-align: top;
                 }
                 .performance-scorecard .table th:first-child,
                 .performance-scorecard .table td:first-child {
                   position: sticky;
                   left: 0;
                   z-index: 1;
                   background-color: #337ab7;
                 }
                 .performance-scorecard .table tr:nth-child(odd) td:first-child,
                 .performance-scorecard .table tr:nth-child(odd) td:nth-child(2) {
                   // background-color: #e9ecef;
                   background-color: #ffffff;
                 }
                 .performance-scorecard .table tr:nth-child(even) td:first-child,
                 .performance-scorecard .table tr:nth-child(even) td:nth-child(2) {
                   background-color: #ffffff;
                 }
                 .performance-scorecard .table tr:nth-child(odd) {
                   // background-color: #e9ecef;
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
                 .performance-scorecard .table .hidden {
                   display: none !important;
                 }
                 .performance-scorecard .table .bdr-rt {
                   border-right: 1px solid #abd0ec;
                 }
                 .performance-scorecard .table thead {
                   position: sticky;
                   top: 0;
                   z-index: 99;
                   font-family: 'FaktPro-Bold';
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
                   background-color: #fff !important;
                   color: #000;
                   text-transform: uppercase;
                   border-bottom: #cfd6de;
                   box-shadow: inset 0 1px 0 0 #cfd6de;
                 }
                 .performance-scorecard .table thead tr:nth-child(2) th :nth-child(1) {
                   background-color: #337ab7 !important;
                   color: #fff;
                 }
                 .performance-scorecard .table thead tr:nth-child(3) {
                   background-color: #fff !important;
                   color: #000;
                   text-transform: uppercase;
                   border-bottom: #cfd6de;
                   box-shadow: inset 0 1px 0 0 #cfd6de;
                 }
                 .performance-scorecard .table thead tr:nth-child(3) th :nth-child(1) {
                   background-color: #337ab7 !important;
                   color: #fff;
                 }
                 .performance-scorecard .table tbody {
                   font-family: 'FaktPro-Normal';
                   font-size: .9rem;
                 }
                 .performance-scorecard .table tbody td {
                   padding: 2px 10px;
                   margin: 0px;
                   border: 1px solid #cfd6de
                 }
                 .performance-scorecard .table tbody tr {
                   border-bottom: 1px solid #37a6f8;
                   border-left: 1px solid #37a6f8
                 }
                 .performance-scorecard .table tbody td:first-child {
                   text-align: start;
                   box-shadow: inset -1px 0 0 0 #cfd6de;
                 }
                 .performance-scorecard .table tbody tr td:not(:first-child) {
                   text-align: right !important;
                 }
                 .performance-scorecard .table tbody .sub-title {
                   font-size: .8rem !important;
                 }
                 .performance-scorecard .table tbody .sub-subtitle{
                   font-size: .7rem !important;
                 }
                 .performance-scorecard .table tbody .text-bold {
                   font-family: 'FaktPro-Bold';
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
      logging: true,
      allowTaint: false,
      useCORS: true,
    };
    if (!div) {
      console.error('Element not found');
      return;
    }
    html2canvas(div, options)
      .then((canvas) => {
        let imgWidth = 286;
        let pageHeight = 204;
        let imgHeight = (canvas.height * imgWidth) / canvas.width;
        let heightLeft = imgHeight;
        const contentDataURL = canvas.toDataURL('image/png');
        let pdfData = new jsPDF('l', 'mm', 'a4', true);
        let position = 5;
        function addExtraDataToPage(pdf: any, extraData: any, positionY: any) {
          pdf.text(extraData, 10, positionY);
        }
        function addPageAndImage(pdf: any, contentDataURL: any, position: any) {
          pdf.addPage();
          pdf.addImage(
            contentDataURL,
            'PNG',
            5,
            position,
            imgWidth,
            imgHeight,
            undefined,
            'FAST'
          );
        }
        pdfData.addImage(
          contentDataURL,
          'PNG',
          5,
          position,
          imgWidth,
          imgHeight,
          undefined,
          'FAST'
        );
        addExtraDataToPage(pdfData, '', position + imgHeight + 10);
        heightLeft -= pageHeight;
        while (heightLeft >= 0) {
          position = heightLeft - imgHeight;
          addPageAndImage(pdfData, contentDataURL, position);
          addExtraDataToPage(pdfData, '', position + imgHeight + 10);
          heightLeft -= pageHeight;
        }
        return pdfData;
      })
      .then((doc) => {
        doc.save('Title Report.pdf');
        // popupWin.close();
        this.spinner.hide();
      });
  }
  sendEmailData(Email: any, notes: any, from: any) {
    this.spinner.show();
    const printContents = document.getElementById('TitleReport')!.innerHTML;
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
            <title>Title Report</title>
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
                 .negative {
                   color: red;
                 }
                 .HideItems{
                   display:none;
                  }            
   .backgrossbtn {
                   float: left;
                   font-size: 11px;
                   background-color: #2d91f1;
                   color: white;
                   padding: 4px;
                   border: 1px solid #2d91f1;
                   cursor: pointer;
                   border-radius: 3px;
               }
                 .Selectedrow {
                   background-color: #d3ecff !important;
                   color: #000 !important;
               }
                 .justify-content-between {
                   justify-content: space-between !important;
               }
               .d-flex {
                   display: flex !important;
               }
                 .bg-white {
                   background: #ffffff !important;
               }
                 .negative {
                   color: red;
                 }
                 .negative {
                   color: red;
                 }
                 .performance-scorecard .table>:not(:first-child) {
                   border-top: 0px solid #ffa51a
                 }
                 .performance-scorecard .table {
                   text-align: center;
                   text-transform: capitalize;
                   border: transparent;
                   width: 100%;
                 }
                 .performance-scorecard .table th,
                 .performance-scorecard .table td{
                   white-space: nowrap;
                   vertical-align: top;
                 }
                 .performance-scorecard .table th:first-child,
                 .performance-scorecard .table td:first-child {
                   position: sticky;
                   left: 0;
                   z-index: 1;
                   background-color: #337ab7;
                 }
                 .performance-scorecard .table tr:nth-child(odd) td:first-child,
                 .performance-scorecard .table tr:nth-child(odd) td:nth-child(2) {
                   // background-color: #e9ecef;
                   background-color: #ffffff;
                 }
                 .performance-scorecard .table tr:nth-child(even) td:first-child,
                 .performance-scorecard .table tr:nth-child(even) td:nth-child(2) {
                   background-color: #ffffff;
                 }
                 .performance-scorecard .table tr:nth-child(odd) {
                   // background-color: #e9ecef;
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
                 .performance-scorecard .table .hidden {
                   display: none !important;
                 }
                 .performance-scorecard .table .bdr-rt {
                   border-right: 1px solid #abd0ec;
                 }
                 .performance-scorecard .table thead {
                   position: sticky;
                   top: 0;
                   z-index: 99;
                   font-family: 'FaktPro-Bold';
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
                   background-color: #fff !important;
                   color: #000;
                   text-transform: uppercase;
                   border-bottom: #cfd6de;
                   box-shadow: inset 0 1px 0 0 #cfd6de;
                 }
                 .performance-scorecard .table thead tr:nth-child(2) th :nth-child(1) {
                   background-color: #337ab7 !important;
                   color: #fff;
                 }
                 .performance-scorecard .table thead tr:nth-child(3) {
                   background-color: #fff !important;
                   color: #000;
                   text-transform: uppercase;
                   border-bottom: #cfd6de;
                   box-shadow: inset 0 1px 0 0 #cfd6de;
                 }
                 .performance-scorecard .table thead tr:nth-child(3) th :nth-child(1) {
                   background-color: #337ab7 !important;
                   color: #fff;
                 }
                 .performance-scorecard .table tbody {
                   font-family: 'FaktPro-Normal';
                   font-size: .9rem;
                 }
                 .performance-scorecard .table tbody td {
                   padding: 2px 10px;
                   margin: 0px;
                   border: 1px solid #cfd6de
                 }
                 .performance-scorecard .table tbody tr {
                   border-bottom: 1px solid #37a6f8;
                   border-left: 1px solid #37a6f8
                 }
                 .performance-scorecard .table tbody td:first-child {
                   text-align: start;
                   box-shadow: inset -1px 0 0 0 #cfd6de;
                 }
                 .performance-scorecard .table tbody tr td:not(:first-child) {
                   text-align: right !important;
                 }
                 .performance-scorecard .table tbody .sub-title {
                   font-size: .8rem !important;
                 }
                 .performance-scorecard .table tbody .sub-subtitle{
                   font-size: .7rem !important;
                 }
                 .performance-scorecard .table tbody .text-bold {
                   font-family: 'FaktPro-Bold';
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
        </html>
    `);
    doc.close();
    const div = doc.getElementById('content');
    if (!div) {
      console.error('Element not found');
      return;
    }
    const options = {
      logging: true,
      allowTaint: false,
      useCORS: true,
      scale: 1 // Adjust scale to fit the page better
    };
    html2canvas(div, options)
      .then((canvas) => {
        let imgWidth = 286;
        let pageHeight = 204;
        let imgHeight = (canvas.height * imgWidth) / canvas.width;
        let heightLeft = imgHeight;
        const contentDataURL = canvas.toDataURL('image/png');
        let pdfData = new jsPDF('l', 'mm', 'a4', true);
        let position = 5;
        pdfData.addImage(
          contentDataURL,
          'PNG',
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
            'PNG',
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
        const pdfFile = this.blobToFile(pdfBlob, 'Title Report.pdf');
        const formData = new FormData();
        formData.append('to_email', Email);
        formData.append('subject', 'Title Report');
        formData.append('file', pdfFile);
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

  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    } else if (value < 0) {
      return false;
    }
    return true;
  }
}
