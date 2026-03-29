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
  imports: [SharedModule, BsDatepickerModule, DateRangePicker],
  standalone: true,
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
})
export class Dashboard {
  NoData: any = false;

  DupFromDate: any = '';
  DupToDate: any = ''
  reportOpenSub!: Subscription;
  reportGetting!: Subscription;
  Pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;
  PreformanceScorecardData: any = []
  CompleteData: any = []
  PreformanceScorecardDataDetails: any = []
  NoScorecardData: any = ''
  LendersIndividual: any = [];
  LendersTotal: any = []
  Lenders: any = []


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

  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent) {
    const clickedInside = (event.target as HTMLElement).closest('.dropdown-toggle, .reportstores-card, .timeframe');
    if (!clickedInside) {
      this.activePopover = -1;
    }
  }
  minDate!: Date;
  maxDate!: Date;
  DateType: any = 'MTD';
  displaytime: any = '';
  FromDate: any = '';
  ToDate: any = '';
  Dates: any = {
    'FromDate': this.FromDate, 'ToDate': this.ToDate, "MaxDate": this.maxDate, 'MinDate': this.minDate, 'DateType': this.DateType, 'DisplayTime': this.displaytime,
    Types: [
      { 'code': 'MTD', 'name': 'MTD' },
      { 'code': 'YTD', 'name': 'YTD' },
      { 'code': 'PYTD', 'name': 'PYTD' },
      { 'code': 'LM', 'name': 'Last Month' },
      { 'code': 'PM', 'name': 'Same Month PY' },
    ]
  }

  constructor(public shared: Sharedservice, public setdates: Setdates, private comm: common, private cp: CurrencyPipe, private toast: ToastService,) {
    this.shared.setTitle(this.comm.titleName + '- Parts Gross Profit Performance Scorecard');


    this.initializeDates('MTD')
    this.maxDate = new Date();
    this.minDate = new Date();
    this.minDate.setFullYear(this.maxDate.getFullYear() - 3);
    this.maxDate.setDate(this.maxDate.getDate());


    this.SetHeaderData();
    this.GetLenders();
  }
  ngOnInit(): void {

  }

  initializeDates(type: any) {
    let dates: any = this.setdates.setDates(type)
    this.FromDate = dates[0];
    this.ToDate = dates[1];
    this.DupFromDate = this.FromDate;
    this.DupToDate = this.ToDate
    localStorage.setItem('time', type);
     this.setDates(this.DateType)
  }
  isDesc: boolean = false;
  column: string = '';
  sort(property: any) {
    this.isDesc = !this.isDesc; //change the direction
    this.column = property;
    let direction = this.isDesc ? 1 : -1;
    // // console.log(property)
    this.PreformanceScorecardData.sort(function (a: any, b: any) {
      if (a[property] < b[property]) {
        return -1 * direction;
      } else if (a[property] > b[property]) {
        return 1 * direction;
      } else {
        return 0;
      }
    });
  }
  GetLenders() {
    console.log(this.FromDate, this.ToDate);
    this.DupFromDate = this.FromDate;
    this.DupToDate = this.ToDate
    this.PreformanceScorecardData = [];
    this.NoData = false;
    this.NoScorecardData = ''
    this.shared.spinner.show();
    const obj = {
      // "fromDate": this.FromDate,
      // "toDate": this.ToDate,
      // 'Stores': this.storeIds.toString()

      "StartDate": this.shared.datePipe.transform(this.FromDate, 'MM-dd-yyyy'),
      "EndDate": this.shared.datePipe.transform(this.ToDate, 'MM-dd-yyyy'),
      "StoreID": ""
    };
    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetPartsScorecard', obj).subscribe(
      (res) => {
        if (res.status == 200) {
          this.PreformanceScorecardData = [];
          if (res.response != undefined) {
            if (res.response.length > 0) {
              // this.LendersTotal = res.response.filter(
              //   (e: any) => e.Location == 'REPORT TOTAL'
              // );
              // this.LendersIndividual = res.response.filter(
              //   (i: any) => i.Location != 'REPORT TOTAL'
              // );
              // this.NoData = false;

              let Data = res.response.reduce((r: any, { GroupName, seq, StoreName, ...rest }: any) => {
                if (!r.some((o: any) => o.seq == seq)) {
                  r.push({
                    GroupName,
                    ...rest,
                    StoreName,
                    seq,
                    subData: res.response.filter((v: any) => v.seq == seq).map((v: any) => ({ ...v, Update: 'N' }))
                    // .sort((a:any, b:any) => a.StoreName.localeCompare(b.StoreName)),
                  });
                }
                return r;
              }, []);


              // this.LendersIndividual.push(this.LendersTotal[0]);
              if (Data && Data.length > 1) {
                this.PreformanceScorecardData = Data;
                let arr = [...Data];
                this.CompleteData = JSON.parse(JSON.stringify(arr));
              } else {
                this.PreformanceScorecardData = []
                this.CompleteData = []
                this.NoScorecardData = 'No Data Found!!'
              }
              console.log(this.PreformanceScorecardData);
              this.shared.spinner.hide();
            } else {
              this.shared.spinner.hide();
              this.NoScorecardData = 'No Data Found!!'

            }
          } else {
            this.shared.spinner.hide();
            this.NoScorecardData = 'No Data Found!!'
          }
        } else {
          this.toast.show(res.status, 'danger', 'Error');
          this.shared.spinner.hide();
          this.NoScorecardData = 'No Data Found!!'
        }
      },
      (error) => {
        this.toast.show('502 Bad Gate Way Error', 'danger', 'Error');
        this.shared.spinner.hide();
        this.NoScorecardData = 'No Data Found!!'
      }
    );
  }


  LMstate: any;
  ngAfterViewInit() {

    this.shared.api.getStores().subscribe((res: any) => {
      if (this.comm.pageName == 'Parts Gross Profit Performance Scorecard') {
        if (res.obj.storesData != undefined) {
           this.groupsArray = res.obj.storesData;
          this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
          this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_name : this.groupName = ''
          this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
          this.getStoresandGroupsValues()
        }
      }
    })

    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res) => {
      if (this.email != undefined) {
        if (res.obj.title == 'Parts Gross Profit Performance Scorecard') {
          if (res.obj.stateEmailPdf == true) {
          //  this.sendEmailData(res.obj.Email, res.obj.notes, res.obj.from);
          }
        }
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res) => {
      if (this.print != undefined) {
        if (res.obj.title == 'Parts Gross Profit Performance Scorecard') {
          if (res.obj.statePrint == true) {
        //    this.GetPrintData();
          }
        }
      }
    });
    this.Pdf = this.shared.api.getExportToPDFAllReports().subscribe((res) => {
      if (this.Pdf != undefined) {
        if (res.obj.title == 'Parts Gross Profit Performance Scorecard') {
          if (res.obj.statePDF == true) {
          //  this.generatePDF();
          }
        }
      }
    });
  
    this.excel = this.shared.api.getExportToExcelAllReports().subscribe((res) => {
      if (this.excel != undefined) {
        this.LMstate = res.obj.state;
        if (res.obj.title == 'Parts Gross Profit Performance Scorecard') {
          if (res.obj.state == true) {
            this.exportToExcel();
          }
        }
      }
    });
  }
  Scrollpercent: any = 0;
  scrollCurrentposition: any = 0
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
  SetHeaderData() {
    const data = {
      title: 'Parts Gross Profit Performance Scorecard',     
      stores: this.storeIds,
      groups: this.groupId,
      fromdate: this.FromDate,
      todate: this.ToDate,
    };
    this.shared.api.SetHeaderData({
      obj: data,
    });
  }
  // GetPrintData() {
  //   window.print();
  // }
  ngOnDestroy() {
    if (this.reportOpenSub != undefined) {
      this.reportOpenSub.unsubscribe()
    }
    if (this.reportGetting != undefined) {
      this.reportGetting.unsubscribe()
    }
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

  grid: any = 'Y'
  RepairData: any = 'Y'
  WarrantyData: any = 'Y'
  ESPData: any = 'Y'
  InternalData: any = 'Y'
  TiresData: any = 'Y'


  toggleGuide(block: any, Cancel: any) {
    if (Cancel == 'X') {
      this.PreformanceScorecardData = this.CompleteData
    }
    if (block == 'C') {
      if (this.RepairData == 'N' && Cancel == 'Y') {
        this.save(block)
      }
      this.RepairData == 'Y' ? this.RepairData = 'N' : this.RepairData = 'Y'

    }
    if (block == 'W') {
      if (this.WarrantyData == 'N' && Cancel == 'Y') {
        this.save(block)
      }
      this.WarrantyData == 'Y' ? this.WarrantyData = 'N' : this.WarrantyData = 'Y'
    }
    if (block == 'E') {
      if (this.ESPData == 'N' && Cancel == 'Y') {
        this.save(block)
      }
      this.ESPData == 'Y' ? this.ESPData = 'N' : this.ESPData = 'Y'
    }
    if (block == 'I') {
      if (this.InternalData == 'N' && Cancel == 'Y') {
        this.save(block)
      }
      this.InternalData == 'Y' ? this.InternalData = 'N' : this.InternalData = 'Y'
    }
    if (block == 'T') {
      if (this.TiresData == 'N' && Cancel == 'Y') {
        this.save(block)
      }
      this.TiresData == 'Y' ? this.TiresData = 'N' : this.TiresData = 'Y'
    }
  }

  save(block: any) {
    let temp: any = []
    this.PreformanceScorecardData.forEach((val: any) => {
      val.subData.forEach((e: any) => {
        if (e.AP_StoreID != 0 && e.Update == 'Y')
          temp.push({
            storeid: e.AP_StoreID,
            paytype: block,
            Target: e[block + '_dup'],
          })
      })
    })
    this.shared.api.postmethod(this.comm.routeEndpoint + 'AddPartsScorecardTargetsGuide', temp).subscribe(
      (res) => {
        if (res.status == 200) {
          this.toast.success('Values Updated Successfully')
          this.GetLenders()
          this.RepairData = 'Y'
          this.WarrantyData = 'Y'
          this.ESPData = 'Y'
          this.InternalData = 'Y'
          this.TiresData = 'Y'
        } else {
          this.toast.show('Something went wrong please try again later', 'danger', 'Error');
        }
      }, (error) => {
        this.toast.show('Something went wrong please try again later', 'danger', 'Error');
      })
  }

  keyPressNumbersincludedecimal(event: KeyboardEvent, data: any): boolean {
    console.log(data);

    data.Update = 'Y';

    const input = event.target as HTMLInputElement;
    const value = input.value ?? '';
    const key = event.key;

    // Allow control/navigation keys
    const ctrlKeys = ['Backspace', 'Tab', 'ArrowLeft', 'ArrowRight', 'Delete', 'Home', 'End'];
    if (ctrlKeys.includes(key)) return true;

    // Allow digits always
    if (key >= '0' && key <= '9') {
      // If there's already a '.', ensure we don't exceed one digit after it
      const dotIndex = value.indexOf('.');
      if (dotIndex !== -1) {
        const selStart = input.selectionStart ?? value.length;
        const selEnd = input.selectionEnd ?? value.length;

        // What will the text look like after this key press?
        const next = value.slice(0, selStart) + key + value.slice(selEnd);
        const nextDotIndex = next.indexOf('.');
        const decimals = nextDotIndex === -1 ? '' : next.slice(nextDotIndex + 1);

        // Block if more than 1 digit appears after the dot
        if (decimals.length > 2) {
          event.preventDefault();
          return false;
        }
      }
      return true;
    }

    // Allow a single '.' (insert '0.' if dot at start)
    if (key === '.') {
      const selStart = input.selectionStart ?? value.length;
      const selEnd = input.selectionEnd ?? value.length;
      const replacing = value.slice(selStart, selEnd);

      // If dot already exists outside the current selection, block
      const valueWithoutSelection = value.slice(0, selStart) + value.slice(selEnd);
      if (valueWithoutSelection.includes('.')) {
        event.preventDefault();
        return false;
      }

      // Optional: auto‑prepend '0' for leading dot
      if ((value.length === 0 || selStart === 0) && replacing !== '.') {
        event.preventDefault();
        const next = '0.' + value.slice(selEnd);
        input.value = next;
        // place caret at end after '0.'
        const pos = 2;
        input.setSelectionRange(pos, pos);
        return false;
      }
      return true;
    }

    // Block everything else
    event.preventDefault();
    return false;
  }

  keyPressNumbers(event: any, data: any) {
    data.Update = 'Y';
    var charCode = event.which ? event.which : event.keyCode;
    // Only Numbers 0-9
    if (charCode < 48 || charCode > 57) {
      event.preventDefault();
      return false;
    } else {
      return true;
    }
  }

  onKeyDown(event: KeyboardEvent, data: any) {
    if (event.key === 'Backspace') {
      console.log('Backspace pressed');
      // Your logic here
      data.Update = 'Y';
    }
  }
  activePopover: number = -1;
 
  togglePopover(popoverIndex: number) {
    if (this.activePopover === popoverIndex) {
      this.activePopover = -1;
    } else {
      this.activePopover = popoverIndex;
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

  viewreport() {
    this.activePopover = -1
    // if (this.storeIds.length == 0) {
    //   this.toast.warning('Please Select Atleast One Store');
    // }
    // else {
    // const data = {
    //   Reference: 'Parts Gross Profit Performance Scorecard',
    //   storeValues: this.storeIds.toString(),
    //   groups: this.groups.toString(),
    // };
    // this.shared.api.SetReports({
    //   obj: data,
    // });
    this.GetLenders()
    // }
  }





  ExcelStoreNames: any = [];
  exportToExcel() {
    let storeNames: any[] = [];
    const store = this.storeIds
    storeNames = this.comm.MobileServiceGL.filter((item: any) =>
      store.includes(item.ac_as_id)
    );
    if (store.length == this.comm.MobileServiceGL.length) {
      this.ExcelStoreNames = 'All Stores'
    } else {
      this.ExcelStoreNames = storeNames.map(function (a: any) {
        return a.storename;
      });
    }
    console.log(store, this.ExcelStoreNames, storeNames);

    // Setup Excel
    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('Parts Gross Profit Performance Scorecard');
    const DATE_EXTENSION = this.shared.datePipe.transform(new Date(), 'MMddyyyy');

    worksheet.views = [{ state: 'frozen', ySplit: 5, topLeftCell: 'A6', showGridLines: false }];

    // Header section (above grid)
    worksheet.addRow([]);
    const titleRow = worksheet.addRow(['Parts Gross Profit Performance Scorecard']);
    titleRow.font = { bold: true, size: 12 };
    worksheet.addRow([]);

    const Stores1 = worksheet.getCell('A3');
    Stores1.value = 'Stores :';
    worksheet.mergeCells('B3', 'Z3');
    const stores1 = worksheet.getCell('B3');
    stores1.value = this.ExcelStoreNames.toString().replaceAll(',', ', ');
    stores1.font = { name: 'Arial', family: 4, size: 9 };
    stores1.alignment = { vertical: 'top', horizontal: 'left', wrapText: true, };

    const Timeframe = worksheet.getCell('A4');
    Timeframe.value = 'Time Frame :';
    worksheet.mergeCells('B4', 'Z4');
    const timeframe = worksheet.getCell('B4');
    timeframe.value = this.shared.datePipe.transform(this.FromDate, 'MM/dd/yyyy') + ' - ' + this.shared.datePipe.transform(this.ToDate, 'MM/dd/yyyy');
    timeframe.font = { name: 'Arial', family: 4, size: 9 };
    timeframe.alignment = { vertical: 'top', horizontal: 'left', wrapText: true, };

    worksheet.addRow('');
    worksheet.getCell('A6');
    worksheet.mergeCells('B6:F6');

    worksheet.getCell('A6').value = this.shared.datePipe.transform(this.FromDate, 'MMMM');
    worksheet.getCell('B6').value = 'Parts Gross Profit Performance Scorecard';

    worksheet.getRow(1).height = 25;

    const PresentYear = this.shared.datePipe.transform(this.FromDate, 'yyyy');
    const FromDate = this.shared.datePipe.transform(this.FromDate, 'dd');
    const ToDate = this.shared.datePipe.transform(this.ToDate, 'dd');
    const PresentMonth = this.shared.datePipe.transform(this.FromDate, 'MMMM');

    ['A6', 'P6'].forEach(key => {
      const cell = worksheet.getCell(key);
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2F5597' } };
    });

    const headers: any = [`${FromDate} - ${ToDate}, ${PresentYear}`, 'Guide', 'Repair Shop', 'Diff', 'Guide', 'Warranty', 'Diff', 'Guide', 'ESP', 'Diff', 'Guide', 'Internal', 'Diff', 'Guide', 'Tires(oth Merch)', 'Diff'];
    const headerRow = worksheet.addRow(headers);
    headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 9 };
    headerRow.eachCell((cell) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '2a91f0' },
      };
      cell.alignment = { horizontal: 'center', vertical: 'middle', indent: 1 };
    });
    // const columnWidths = new Array(headers.length).fill(10);


    const columnWidths = new Array(6).fill(10);
    // Main Data Rows
    this.PreformanceScorecardData.forEach((item: any, index: number) => {

      item.subData.forEach((sub: any, j: any) => {
        if (sub.AP_StoreID != 0) {
          const subData = [
            sub.StoreName ? sub.StoreName : '',
            sub['C'] ? sub['C'] + '%' : '',
            sub['Repair Shop'] ? sub['Repair Shop'] + '%' : '',
            sub['C_Diff'] ? sub['C_Diff'] + '%' : '',
            sub['W'] ? sub['W'] + '%' : '',
            sub['Warranty Service'] ? sub['Warranty Service'] + '%' : '',
            sub['W_Diff'] ? sub['W_Diff'] + '%' : '',
            sub['E'] ? sub['E'] + '%' : '',
            sub['Esp Repair Shop'] ? sub['Esp Repair Shop'] + '%' : '',
            sub['E_Diff'] ? sub['E_Diff'] + '%' : '',
            sub['I'] ? sub['I'] + '%' : '',
            sub['Internal Service'] ? sub['Internal Service'] + '%' : '',
            sub['I_Diff'] ? sub['I_Diff'] + '%' : '',
            sub['T'] ? sub['T'] + '%' : '',
            sub.Tires ? sub.Tires + '%' : '',
            sub['T_Diff'] ? sub['T_Diff'] + '%' : '',

          ];
          const subRow = worksheet.addRow(subData);
          subRow.font = { size: 9 };
          subRow.getCell(4).alignment = { horizontal: 'right', indent: 1 };
          subRow.getCell(5).alignment = { horizontal: 'right', indent: 1 };
          subRow.getCell(6).alignment = { horizontal: 'right', indent: 1 };
          subRow.eachCell((cell) => {
            // cell.fill = {
            //   type: 'pattern',
            //   pattern: 'solid',
            //   fgColor: { argb: 'd2deed' },
            // };
            cell.alignment = { ...cell.alignment, indent: 1 };
          });

          subData.forEach((val, i) => {
            const length = val?.toString().length || 0;
            columnWidths[i] = Math.max(columnWidths[i], length);
          });
        }
      });

      const mainData = [
        item.GroupName ? item.GroupName : '',
        item['C'] ? item['C'] + '%' : '',
        item['Repair Shop'] ? item['Repair Shop'] + '%' : '',
        item['C_Diff'] ? item['C_Diff'] + '%' : '',
        item['W'] ? item['W'] + '%' : '',
        item['Warranty Service'] ? item['Warranty Service'] + '%' : '',
        item['W_Diff'] ? item['W_Diff'] + '%' : '',
        item['E'] ? item['E'] + '%' : '',
        item['Esp Repair Shop'] ? item['Esp Repair Shop'] + '%' : '',
        item['E_Diff'] ? item['E_Diff'] + '%' : '',
        item['I'] ? item['I'] + '%' : '',
        item['Internal Service'] ? item['Internal Service'] + '%' : '',
        item['I_Diff'] ? item['I_Diff'] + '%' : '',
        item['T'] ? item['T'] + '%' : '',
        item.Tires ? item.Tires + '%' : '',
        item['T_Diff'] ? item['T_Diff'] + '%' : '',

      ];
      const mainRow = worksheet.addRow(mainData);
      mainRow.font = { size: 9 };
      mainRow.getCell(4).alignment = { horizontal: 'right', indent: 1 };
      mainRow.getCell(5).alignment = { horizontal: 'right', indent: 1 };
      mainRow.getCell(6).alignment = { horizontal: 'right', indent: 1 };
      mainRow.eachCell((cell) => {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'd2deed' },
        };
        cell.alignment = { ...cell.alignment, indent: 1 };
      });

      mainData.forEach((val, i) => {
        const length = val?.toString().length || 0;
        columnWidths[i] = Math.max(columnWidths[i], length);
      });

      columnWidths.forEach((width, i) => {
        worksheet.getColumn(i + 1).width = Math.max(15, width + 2);
      });
    })

    // Set Column Widths


    // Export to Excel File
 
      this.shared.exportToExcel(workbook, `Parts Gross Profit Performance Scorecard_${DATE_EXTENSION}.xlsx`);
   
  }

  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    } else if (value < 0) {
      return false;
    }
    return true;
  }


}
