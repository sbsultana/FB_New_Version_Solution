import { Component, Injector, HostListener } from '@angular/core';
import { Api } from '../../../../Core/Providers/Api/api';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
import {  Subscription } from 'rxjs';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { Workbook } from 'exceljs';
import * as FileSaver from 'file-saver';
import { DatePipe } from '@angular/common';
import { NgxSpinnerService } from 'ngx-spinner';

const EXCEL_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
const EXCEL_EXTENSION = '.xlsx';
@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [SharedModule],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
})
export class Dashboard {
  Partsdetailsdata: any = [];
  AccountingMappingData: any = [];
  stores: any = [];
  NoData: any = false;
  NoDatainitail: any = true;
  PageCount = 1;
  Pagination: boolean = false;
  LastCount!: boolean;
  LMstate: any;
  selectedStore = '0';
  selectedstrid = 0;
  searchText: any = '';
  storeIds: any = '0';
  groups: any = 1;
  reportOpenSub!: Subscription;
  reportGetting!: Subscription;
  excel!: Subscription;
  constructor(
    private datepipe: DatePipe,
    private toast: ToastService,
    public apiSrvc: Api,
    public shared: Sharedservice,
    private spinner: NgxSpinnerService,
  ) {
    this.NoData = false;
    this.NoDatainitail = true
    if (localStorage.getItem('Fav') != 'Y') {
      const data = {
        title: 'Search Parts',
        path1: '',
        path2: '',
        path3: '',
        stores: this.storeIds,
        groups: this.groups,
        count: 0
      };
      this.apiSrvc.SetHeaderData({
        obj: data,
      });
    } else {

    }
  }

  getStores() {
    const obj = {
      id: '1',
      userid: JSON.parse(localStorage.getItem('UserDetails')!).userid
    }
    this.apiSrvc.postmethodOne(this.shared.common.routeEndpoint + 'GetStoresbyGroupuserid', obj).subscribe((res: any) => {
      this.stores = res.response;
    })
   
    if (this.stores.length > 0) {
      this.selectedStore = '0';
      this.selectedstrid = 0;
    } else {
     
    }

  }

  getPartsdetailsdata() {
    this.Partsdetailsdata = [];
    this.NoData = false;
    this.spinner.show();
    const obj = {

      "partnumber": this.searchText,


    };
    const curl = this.shared.common.routeEndpoint + 'GetPartsInventoryData';

    this.apiSrvc.postmethod(curl, obj).subscribe((res) => {
      const currentTitle = document.title;
      this.apiSrvc.logSaving(curl, {}, '', res.message, currentTitle);
      if (res.status == 200) {
        if (res.response != undefined) {
          if (res.response.length > 0) {
            //////console.log(res.response.length,  this.Partsdetailsdata);
            this.Partsdetailsdata = res.response;
            this.spinner.hide();
            this.NoData = false;
            this.NoDatainitail = false;
          }
          else {
            this.spinner.hide();
            this.NoData = true;
            this.NoDatainitail = false;

          }
        }
        else {
          // this.toast.error('Empty Response');
          this.NoData = true;
          this.NoDatainitail = false;
          this.spinner.hide();
        }
      } else {
        // this.toast.error(res.status, '');
        this.spinner.hide();
        this.NoData = true;
        this.NoDatainitail = false;
      }
    },
      (error) => {
        this.toast.show('502 Bad Gate Way Error');
        this.spinner.hide();
        this.NoData = true;
        this.NoDatainitail = false;
      }
    );

  }
  index = '';
  commentobj = {};

  close(data: any) {
    console.log(data);
    this.index = '';
  }
  addcmt(data: any) {
    if (data == 'AD') {
      // this.getAccountingMappingTabALL();
    }
  }

  commentopen(item: any, i: any) {
    console.log(item);

    this.index = i.toString();
    this.commentobj = {
      TYPE: item.data1,
      NAME: item.data1,
      STORES: item.data1,
      STORENAME: item.data1,
      Month: '',
      ModuleId: '10',
      ModuleRef: 'AM',
      state: 1,
      indexval: i,
    };
  }
  clearFilter() {
    this.searchText = '';
    // this.searchFilter();
    this.Partsdetailsdata = []

    this.AccountingMappingData = []

    this.NoData = false
    this.NoDatainitail = true
  }

  searchFilter() {
    if (this.searchText != '') {
      this.getPartsdetailsdata();
    } else {
      this.NoData = false;
      this.NoDatainitail = true;
      this.Partsdetailsdata = []

      this.AccountingMappingData = []


    }
  }
  onKeypressEvent(event: any) {
    console.log(event.target.value);
    if (event.keyCode == 13) {
      if (this.searchText != '') {
        this.getPartsdetailsdata();
      } else {
        this.Partsdetailsdata = []
        this.AccountingMappingData = []
        this.NoData = false;
        this.NoDatainitail = true;
      }
    }
  }


  ngAfterViewInit(): void {
    this.reportGetting = this.apiSrvc.GetReports().subscribe((data) => {
      if (this.reportGetting != undefined) {
        //////console.log(data)
        if (data.obj.Reference == 'Search Parts') {
          // this.Filter = data.obj.Filter;
          const headerdata = {
            title: 'Search Parts',
            path1: '',
            path2: '',
            path3: '',
            // filter: this.Filter,
            stores: data.obj.storeValues,
            groups: data.obj.groups
          };

          this.apiSrvc.SetHeaderData({
            obj: headerdata,
          });
          this.getPartsdetailsdata();
        }
      }
    });
    this.excel = this.apiSrvc.getExportToExcelAllReports().subscribe((res) => {
      if (this.excel != undefined) {
        this.LMstate = res.obj.state;
        if (res.obj.title == 'Search Parts') {
          if (res.obj.state == true) {
            this.exportToExcel();
          }
        }
      }
    });

  }
  ExcelStoreNames: any = [];
  exportToExcel() {
    let storeNames: any = [];
    let store = this.storeIds.split(',');
    const obj = {
      id: this.groups,
      userid: JSON.parse(localStorage.getItem('UserDetails')!).userid,
    };
    this.apiSrvc
      .postmethodOne(this.shared.common.routeEndpoint + 'GetStoresbyGroupuserid', obj)
      .subscribe((res: any) => {
        storeNames = res.response.filter((item: any) =>
          store.some((cat: any) => cat === item.ID.toString())
        );
        // console.log(store, res.response.length);

        if (store.length == res.response.length) {
          this.ExcelStoreNames = 'All Stores'
        } else {
          this.ExcelStoreNames = storeNames.map(function (a: any) {
            return a.storename;
          });
        }
        const Partsdetailsdata = this.Partsdetailsdata.map((_arrayElement: any) =>
          Object.assign({}, _arrayElement)
        );
        const workbook = new Workbook();
        const worksheet = workbook.addWorksheet('Search Parts');
        worksheet.views = [
          {
            state: 'frozen',
            ySplit: 11, // Number of rows to freeze (2 means the first two rows are frozen)
            topLeftCell: 'A12', // Specify the cell to start freezing from (in this case, the third row)
            showGridLines: false,
          },
        ];
        worksheet.addRow('');
        const titleRow = worksheet.addRow(['Search Parts']);
        titleRow.eachCell((cell, number) => {
          cell.alignment = { indent: 1, vertical: 'top', horizontal: 'left' };
        });
        titleRow.font = { name: 'Arial', family: 4, size: 12, bold: true };
        worksheet.addRow('');

        const DateToday = this.datepipe.transform(
          new Date(),
          'MM/dd/yyyy h:mm:ss a'
        );
        const DATE_EXTENSION = this.datepipe.transform(
          new Date(),
          'MMddyyyy'
        );
        worksheet.addRow([DateToday]).font = { name: 'Arial', family: 4, size: 9 };

        const ReportFilter = worksheet.addRow(['Report Controls :']);
        ReportFilter.font = { name: 'Arial', family: 4, size: 10, bold: true };

        const Groups = worksheet.getCell('A6');
        Groups.value = 'Group :';
        Groups.font = { name: 'Arial', family: 4, size: 9, bold: true };
        const groups = worksheet.getCell('B6');
        groups.value =
          this.groups == 1 ? this.shared.common.excelName : (this.groups == 2 ? 'Domestic' : (this.groups == 3 ? 'Import' : (this.groups == 4 ? 'Warehouse' : '-')));
        groups.font = { name: 'Arial', family: 4, size: 9 };
        groups.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
        worksheet.mergeCells('B7', 'K9');
        const Stores = worksheet.getCell('A7');
        Stores.value = 'Stores :'
        const stores = worksheet.getCell('B7');
        stores.value = this.ExcelStoreNames == 0
          ? 'All Stores'
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

        worksheet.addRow('');

        let Headings = [
          'Dealer Name',
          'Class',
          'Comment',
          'Comp',
          'Cost',
          'Date Added',
          'Description',
          'Exchange',
          'List',
          'Last Count Date',
          'Last Sale Date',
          'Last Trans Date',
          'Memo Part Flag',
          'MFG controlled',
          'Misc',
          'Multiple',
          'New Order Quantity',
          'Onhand',
          'On Order Quantity',
          'Pack',
          'Pack Multiple Flag',
          'Price Unit',
          'Manufacturer',
          'Part Number',
          'Sortkey1',
          'Source',
          'Special Status',
          'Trade',
          'Unit Weight',
          'Update Code'
        ];
        const headerRow = worksheet.addRow(Headings);
        headerRow.font = {
          name: 'Arial',
          family: 4,
          size: 8,
          bold: true,
          color: { argb: 'FFFFFF' },
        };
        headerRow.alignment = { indent: 1, vertical: 'top', horizontal: 'center' };
        headerRow.eachCell((cell, number) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: '2a91f0' },
            bgColor: { argb: 'FF0000FF' },
          };
          cell.border = { right: { style: 'dotted' } };
          cell.alignment = { vertical: 'top', horizontal: 'center' };
        });

        for (const d of Partsdetailsdata) {
          const Data1 = worksheet.addRow([
            d.dealername == '' ? '-' : d.dealername == null ? '-' : d.dealername,
            d.class == '' ? '-' : d.class == null ? '-' : d.class,
            d.comment == '' ? '-' : d.comment == null ? '-' : d.comment,
            d.comp == '' ? '-' : d.comp == null ? '-' : d.comp,
            d.dateadded == '' ? '-' : d.dateadded == null ? '-' : d.dateadded,
            d.description == '' ? '-' : d.description == null ? '-' : d.description,
            d.exchange == '' ? '-' : d.exchange == null ? '-' : d.exchange,
            d.list == '' ? '-' : d.list == null ? '-' : d.list,
            d.lastcountdate == '' ? '-' : d.lastcountdate == null ? '-' : d.lastcountdate,
            d.lastsaledate == '' ? '-' : d.lastsaledate == null ? '-' : d.lastsaledate,
            d.lasttransdate == '' ? '-' : d.lasttransdate == null ? '-' : d.lasttransdate,
            d.memopartflag == '' ? '-' : d.memopartflag == null ? '-' : d.memopartflag,
            d.mfgcontrolled == '' ? '-' : d.mfgcontrolled == null ? '-' : d.mfgcontrolled,
            d.misc == '' ? '-' : d.misc == null ? '-' : d.misc,
            d.multiple == '' ? '-' : d.multiple == null ? '-' : d.multiple,
            d.neworderqty == '' ? '-' : d.neworderqty == null ? '-' : d.neworderqty,
            d.onorderqty == '' ? '-' : d.onorderqty == null ? '-' : d.onorderqty,
            d.pack == '' ? '-' : d.pack == null ? '-' : d.pack,
            d.packmultipleflag == '' ? '-' : d.packmultipleflag == null ? '-' : d.packmultipleflag,
            d.priceunit == '' ? '-' : d.priceunit == null ? '-' : d.priceunit,
            d.manufacturer == '' ? '-' : d.manufacturer == null ? '-' : d.manufacturer,
            d.partnumber == '' ? '-' : d.partnumber == null ? '-' : d.partnumber,
            d.sortkey1 == '' ? '-' : d.sortkey1 == null ? '-' : d.sortkey1,
            d.source == '' ? '-' : d.source == null ? '-' : d.source,
            d.specialstatus == '' ? '-' : d.specialstatus == null ? '-' : d.specialstatus,
            d.trade == '' ? '-' : d.trade == null ? '-' : d.trade,
            d.sortkey1 == '' ? '-' : d.sortkey1 == null ? '-' : d.sortkey1,
            d.source == '' ? '-' : d.source == null ? '-' : d.source,
            d.unitweight == '' ? '-' : d.unitweight == null ? '-' : d.unitweight,
            d.updatecode == '' ? '-' : d.updatecode == null ? '-' : d.updatecode,
          ]);
          // Data1.outlineLevel = 1; // Grouping level 1
          Data1.font = { name: 'Arial', family: 4, size: 8 };
          Data1.alignment = { vertical: 'top', horizontal: 'center' };
          // Data1.getCell(1).alignment = {indent: 1,vertical: 'top', horizontal: 'center'}
          Data1.eachCell((cell, number) => {
            cell.border = { right: { style: 'dotted' } };
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
        worksheet.getColumn(1).width = 30;
        worksheet.getColumn(2).width = 10;
        worksheet.getColumn(3).width = 10;
        worksheet.getColumn(4).width = 10;
        worksheet.getColumn(5).width = 25;
        worksheet.getColumn(6).width = 20;
        worksheet.getColumn(7).width = 10;
        worksheet.getColumn(8).width = 10;
        worksheet.getColumn(9).width = 25;
        worksheet.getColumn(10).width = 25;
        worksheet.getColumn(11).width = 25;
        worksheet.getColumn(12).width = 10;
        worksheet.getColumn(13).width = 10;
        worksheet.getColumn(14).width = 10;
        worksheet.getColumn(15).width = 10;
        worksheet.getColumn(16).width = 10;
        worksheet.getColumn(17).width = 10;
        worksheet.getColumn(18).width = 10;
        worksheet.getColumn(19).width = 15;
        worksheet.getColumn(20).width = 10;
        worksheet.getColumn(21).width = 15;
        worksheet.getColumn(22).width = 10;
        worksheet.getColumn(23).width = 15;
        worksheet.getColumn(24).width = 15;
        worksheet.getColumn(25).width = 10;
        worksheet.getColumn(26).width = 10;
        worksheet.getColumn(27).width = 20;
        worksheet.getColumn(28).width = 10;
        worksheet.getColumn(29).width = 10;
        worksheet.getColumn(30).width = 10;

        worksheet.addRow([]);
        workbook.xlsx.writeBuffer().then((data: any) => {
          const blob = new Blob([data], {
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          });
          FileSaver.saveAs(blob, 'Parts Detials_' + DATE_EXTENSION + EXCEL_EXTENSION);
        });
      });
  }

}
