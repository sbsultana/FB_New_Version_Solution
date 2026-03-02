import { Component, ElementRef, HostListener, Renderer2 } from '@angular/core';
import {
  BsDatepickerConfig,
  BsDatepickerModule,
} from 'ngx-bootstrap/datepicker';
import { NgxSpinnerService } from 'ngx-spinner';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { CurrencyPipe, DatePipe, DecimalPipe } from '@angular/common';
import { common } from '../../../../common';
import { Title } from '@angular/platform-browser';
import { NgbActiveModal, NgbModal } from '@ng-bootstrap/ng-bootstrap';
import { Api } from '../../../../Core/Providers/Api/api';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { environment } from '../../../../../environments/environment';
import { Subscription } from 'rxjs';
import { Stores } from '../../../../CommonFilters/stores/stores';
import numeral from 'numeral';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';

@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [SharedModule, BsDatepickerModule, Stores],
  providers: [CurrencyPipe, DatePipe, DecimalPipe],
  templateUrl: './dashboard.html',
  styleUrls: ['./dashboard.scss']
})
export class Dashboard {
  dynamicTitle = 'Financial Statement';
  FSData: any[] = [];
  Month: string | null = null;
  // storeIds: any = '';
  storeName: any;
  presentDate: Date | null = null;
  lastFourMonthsArray: string[] = [];
  FsNegativeData: any[] = [];
  FinancialStatementExcel: any[] = [];
  NodataFound = false;
  NoData = false;
  excel!: Subscription;
  selectedDate: Date = new Date();
  currentMonth!: Date;
  // groupId: any = [1];
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
  activePopover: number = -1;
  ngChanges: any;

  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent) {
    const clickedInside = (event.target as HTMLElement).closest('.dropdown-toggle, .reportstores-card , .timeframe');
    if (!clickedInside) {
      this.activePopover = -1;
    }
  }

  constructor(
    public apiSrvc: Api,
    public shared: Sharedservice,
    private spinner: NgxSpinnerService,
    private ngbmodal: NgbModal,
    private ngbmodalActive: NgbActiveModal,
    private renderer: Renderer2,
    private el: ElementRef,
    private title: Title,
    private datepipe: DatePipe,
    private currencyPipe: CurrencyPipe,
    private decimalPipe: DecimalPipe,
    private toast: ToastService,
    private comm: common
  ) {
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

    this.shared.setTitle(this.shared.common.titleName + '-Financial Statement');
    // if (localStorage.getItem('Fav') !== 'Y') {
    const data = { title: this.dynamicTitle, stores: '2' };
    this.shared.api.SetHeaderData({ obj: data });
    const lastMonth = new Date();
    let today = new Date();
    if (today.getDate() < 7) {
      this.selectedDate = new Date(lastMonth.setMonth(lastMonth.getMonth() - 1));
    } else {
      this.selectedDate = new Date(lastMonth.setMonth(lastMonth.getMonth()));
    }
    this.currentMonth = this.selectedDate;
    this.onMonthChange(this.currentMonth);
    // }
  }



  bsConfig: Partial<BsDatepickerConfig> = {
    dateInputFormat: 'MMMM/YYYY',
    minMode: 'month',
    maxDate: new Date()
  };

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
      this.currentMonth = this.selectedDate;
      this.onMonthChange(this.currentMonth);
    }
  }

  onMonthChange(monthDate: Date) {
    this.spinner.show();
    this.lastFourMonthsArray = [];
    if (!monthDate) return;
    this.selectedDate = monthDate;
    this.lastFourMonthsArray = [];
    for (let i = 1; i <= 3; i++) {
      const lastMonthDate = new Date(monthDate.getFullYear(), monthDate.getMonth() - i, 1);
      const formattedDate = lastMonthDate.toLocaleDateString('en-US', {
        year: 'numeric',
        month: 'long'
      });
      this.lastFourMonthsArray.push(formattedDate);
    }
    const dateFormatted = this.datepipe.transform(monthDate, 'dd-MMM-yyyy');
    const obj = {
      stores: this.storeIds,
      DATE: dateFormatted,
    };
    const curl = environment.apiUrl + this.comm.routeEndpoint + 'GetFinStatementReport';
    this.apiSrvc.postmethod(this.comm.routeEndpoint + 'GetFinStatementReport', obj).subscribe({
      next: (res) => {
        const currentTitle = document.title;
        this.apiSrvc.logSaving(curl, {}, '', res.message, currentTitle);
        this.spinner.hide();

        if (res.status !== 200) {
        
          this.toast.show('Invalid Details', 'danger', 'Error');
          return;
        }

        this.FSData = res.response || [];
        this.FinancialStatementExcel = this.FSData;
        this.FsNegativeData = this.FSData;
        this.FSData = this.FSData.map((v: any) => ({
          ...v,
          Fs_Titles: 'SubTitle',
          ValueType: '$',
        }));
        this.FSData.forEach((val: any) => {
          const sentence = val.DisplayName || '';
          if (sentence.match(/Units?|Retail/i)) val.ValueType = '#';
          else if (sentence.match(/Percent|Absorption/i)) val.ValueType = '%';
          else val.ValueType = '$';
        });
        this.FSData.forEach((val: any) => {
          const label = val.LabelName;
          if (
            [
              'TotalSellingGross-New', 'SellingGross%VariableGross-New', 'SellingGrossPUS-New',
              'TotalSellingGross-Used', 'SellingGross%VariableGross-Used', 'SellingGrossPUS-Used',
              'Sales-Fixed', 'Gross-Fixed', 'Gross%Sales-Fixed', 'Gross%TotalDeptExpIncludingFixedExp-Fixed',
              'SellingGross-AllDepts', 'SellingGross%Gross-AllDepts', 'TotalFixedExpenses',
              'FixedExpenses%Gross-AllDepts', 'OpertingProfit-AllDepts', 'OpertingProfit%Gross-AllDepts',
              'DealerSalary', 'DealerSalary%Gross-AllDepts', 'OtherAdditions&Reductions',
              'OtherAddDeducts%Gross-AllDepts', 'NetIncome-BeforeTaxes',
              'NetIncomeBeforeTaxes%Gross-AllDepts', 'Payroll-AllDepts', 'TotalDealershipSales',
              'Payroll%Gross-AllDepts', 'SellingGrossPercent', 'SellingGross-Service',
              'SellingGross%DeptGross-Service', 'SellingGross-Parts', 'SellingGross%DeptGross-Parts',
            ].includes(label)
          ) {
            val.Fs_Titles = 'Title';
          } else if (
            ['TotalGrossProfit', 'SellingGross', 'Payroll'].includes(label)
          ) {
            val.Fs_Titles = 'TitleCenter';
          }
        });
        if (this.FSData.length > 0) {
          this.FSData.forEach((x: any) => {
            (x.SubData = []), (x.data2sign = '+');
          });
          this.NoData = false;
        } else {
          this.NoData = true;
        }
        console.log('FS Data', this.FSData);
        this.NodataFound = true;
        this.NoData = this.FSData.length === 0;
      },
      error: (err) => {
        console.error(err);
        this.spinner.hide();
      },
    });
  }
  ngAfterViewInit() {

    this.shared.api.getStores().subscribe((res: any) => {
      if (this.shared.common.pageName == 'Financial Statement') {
        if (res.obj.storesData != undefined) {
          this.groupsArray = res.obj.storesData;
          this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
          this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_name : this.groupName = ''
          this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
          this.getStoresandGroupsValues()
        }
      }
    })
    this.excel = this.shared.api.getExportToExcelAllReports().subscribe((res: { obj: { title: string; state: boolean; }; }) => {
      if (this.excel != undefined) {
        if (res.obj.title == 'Financial Statement') {
          if (res.obj.state == true) {
            this.exportToExcel();
          }
        }
      }
    });


  }
  // @HostListener('document:click', ['$event'])
  // onDocumentClick(event: MouseEvent) {
  //   const clickedInside = (event.target as HTMLElement).closest('.dropdown-toggle, .reportstores-card, .timeframe');
  //   if (!clickedInside) {
  //     this.activePopover = -1;
  //   }
  // }

  togglePopover(popoverIndex: number) {
    this.activePopover = this.activePopover === popoverIndex ? -1 : popoverIndex;
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
  expandorcollapse(ind: any, e: any, ref: any, Item: any, Department: any) {
    let id = (e.target as Element).id;
    console.log(ref, Item);
    if (id == 'D_' + ind) {
      console.log(id);
      if (ref == '-') {
        this.FSData[ind].data2sign = '+';
        this.FSData[ind].SubData = [];
      }
      if (ref == '+') {
        this.FSData[ind].data2sign = '-';
        this.GetFinancialStatementLayerOne(Item, ind, Department);
      }
    }
    if (id == 'DN_' + ind) {
      if (ref == '-') {
        this.FinancialStatementLayerOne[ind].dataLayer2sign = '+';
        this.FinancialStatementLayerOne[ind].SubLayerData = [];
      }
      if (ref == '+') {
        this.FinancialStatementLayerOne[ind].dataLayer2sign = '-';
        // this.GetFixedExpenseLayerTwo(parentData, Item, ind);
        this.FSData[ind].data2sign = '-';
      }
    }
  }
  FinancialStatementLayerOne: any;
  FinancialStatementLayerOneKeys: any;
  GetFinancialStatementLayerOne(Object: any, index: number, Department: any): void {
    this.spinner.show();
    this.FinancialStatementLayerOne = [];
    this.FinancialStatementLayerOneKeys = [];
    const DateToday = this.datepipe.transform(
      new Date(this.selectedDate),
      'dd-MMM-yyyy'
    );
    const obj = {
      "Store": this.storeIds,
      "Department": Object.Department,
      "DATE": DateToday,
      "Headname": Object.DisplayName
    };

    this.apiSrvc
      .postmethod(this.comm.routeEndpoint + 'GetFinStatementExpenseReportDetail', obj)
      .subscribe(
        (x: any) => {
          if (x.status === 200) {
            this.FinancialStatementLayerOne = x.response;
            this.FinancialStatementLayerOne = this.FinancialStatementLayerOne.map((v: any) => ({
              ...v,
              Fs_Titles: 'SubTitle',
              ValueType: '$'
            }));
            this.FinancialStatementLayerOne.forEach((val: any) => {
              const sentence: string = val['SUBTYPE DETAIL'];
              const Units: string = 'Units';
              const Percentage: string = 'Percent';
              const index: number = sentence.indexOf(Units);
              const index2: number = sentence.indexOf(Percentage);
              if (index !== -1) {
                val.ValueType = '#'
              } else if (index2 !== -1) {
                val.ValueType = '%'
              } else {
                val.ValueType = '$'
              }
            })
            console.log('GetFinStatementExpenseDetail', this.FinancialStatementLayerOne);

            if (
              this.FinancialStatementLayerOne &&
              this.FinancialStatementLayerOne.length > 0
            ) {
              // this.FinancialStatementLayerOne.forEach((item: any) => {
              //   item.SubLayerData = [];
              //   item.dataLayer2sign = '+';
              // });
              this.FSData.forEach((val: any) => {
                val.SubData = [];
                val.data2sign = '+';
                this.FinancialStatementLayerOne.forEach((ele: any) => {
                  console.log(ele.Headname);
                  if (val.DisplayName.replace(/\s+/g, '') === ele.Headname && val.Department === Department) {
                    console.log(ele.Headname);
                    val.SubData.push(ele);
                    val.data2sign = '-';
                  }
                });
              });
              this.NoData = false;
            } else {
              this.NoData = false;
            }
            this.spinner.hide();
          }
        },
        (error: any) => {
          console.error(error);
          this.spinner.hide();
        }
      );
  }

  isSpecialRow(name: string): boolean {
    return [
      'Total Selling Gross (New)',
      'Total Selling Gross (Used)',
      'Net Income Before Taxes',
      'Total Operating Department Profit'
    ].includes(name);
  }

  getStyle(value: any, valueType: string) {
    if (value === null || value === '' || value === 0 || value === '0') {
      return { 'justify-content': 'center' };
    }
    switch (valueType) {
      case '$':
        return { 'justify-content': 'space-between' };
      case '#':
        return { 'justify-content': 'flex-end' };
      case '%':
        return { 'justify-content': 'center' };
      default:
        return { 'justify-content': 'initial' };
    }
  }
  formatValue(value: any, valueType: string): string {
    if (value === null || value === undefined || value === 0 || value === '0') {
      return '-';
    } else if (valueType === '#') {
      return typeof value === 'number' ? numeral(value).format('0,0') : value;
    } else if (valueType === '$') {
      return typeof value === 'number' ? numeral(value).format('0,0') : value;
    } else if (valueType === '%') {
      return typeof value === 'number' ? value.toLocaleString() + '%' : value;
    } else {
      return value;
    }
  }
  index: string = '';
  selectRow(i: number) {
    this.index = i.toString();
  }

  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    }
    return false;
  }
  // exportToExcel(): void {
  //   const workbook = this.shared.getWorkbook();
  //   const worksheet = workbook.addWorksheet('Financial Statement');

  //   const formattedMonth = this.shared.datePipe.transform(this.Month, 'MMMM yyyy');


  //   const currentMonth = this.shared.datePipe.transform(this.Month, 'MMMM yyyy');
  //   const storeName = 'WesternAuto';

  //   worksheet.mergeCells('A1:A1');
  //   worksheet.mergeCells('B1:G1');
  //   worksheet.getCell('A1').value = currentMonth;
  //   worksheet.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' };
  //   worksheet.getCell('A1').font = { bold: true, color: { argb: 'FFFFFFFF' } };

  //   worksheet.getCell('B1').value = storeName;
  //   worksheet.getCell('B1').alignment = { horizontal: 'center', vertical: 'middle' };
  //   worksheet.getCell('B1').font = { bold: true, color: { argb: '#000' } };

  //   worksheet.getRow(1).fill = {
  //     type: 'pattern',
  //     pattern: 'solid',
  //     fgColor: { argb: 'FFFFFFFF' },
  //   };
  //   worksheet.getRow(1).height = 25;


  //   const dynamicMonths = this.lastFourMonthsArray.map((m: any) =>
  //     this.shared.datePipe.transform(m, 'MMMM yyyy')
  //   );

  //   const headerRow = ['Department', 'Actual', 'Budget', 'Variance', ...dynamicMonths];
  //   worksheet.addRow(headerRow);

  //   const headerRowIndex = worksheet.lastRow?.number ?? 2;
  //   const headerRowRef = worksheet.getRow(headerRowIndex);
  //   headerRowRef.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  //   headerRowRef.alignment = { horizontal: 'center', vertical: 'middle' };
  //   headerRowRef.fill = {
  //     type: 'pattern',
  //     pattern: 'solid',
  //     fgColor: { argb: 'FF2F5597' },
  //   };
  //   headerRowRef.height = 22;



  // this.FSData.forEach((data: any) => {

  //   const mainRowValues = [
  //     data.DisplayName,
  //     data.Actual,
  //     data.Budget,
  //     data.Variance,
  //     data.LM1,
  //     data.LM2,
  //     data.LM3
  //   ];
  //   const mainRow = worksheet.addRow(mainRowValues);


  //   if (data.IsHeader === 'T') {
  //     mainRow.font = { bold: true };
  //     mainRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDDEBF7' } };
  //   } else if (data.IsHeader === 'Y') {
  //     mainRow.font = { bold: true };
  //     mainRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFBFD4F2' } };
  //   }

  //   mainRow.eachCell((cell, colNum) => {
  //     if (colNum === 1) {
  //       cell.alignment = { horizontal: 'left', vertical: 'middle' };
  //     } else {
  //       const val = cell.value;
  //       if (val == null || val === '-') {
  //         cell.value = '-';
  //       } else if (typeof val === 'number') {
  //         if (val < 0)
  //           cell.value = `($${Math.abs(val).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })})`;
  //         else
  //           cell.value = `$${val.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
  //       }
  //       cell.alignment = { horizontal: 'right', vertical: 'middle' };
  //     }
  //   });


  //   if (data.SubData && data.SubData.length > 0) {
  //     data.SubData.forEach((sub: any) => {
  //       const subRowValues = [
  //         `   ${sub['SUBTYPE DETAIL'] ?? ''}`, // indent the name
  //         sub.Actual,
  //         sub.Budget,
  //         sub.Variance,
  //         sub.LM1,
  //         sub.LM2,
  //         sub.LM3
  //       ];
  //       const subRow = worksheet.addRow(subRowValues);


  //       subRow.fill = {
  //         type: 'pattern',
  //         pattern: 'solid',
  //         fgColor: { argb: 'FFE8EEF7' },
  //       };

  //       subRow.font = { size: 10, color: { argb: 'FF000000' } };

  //       subRow.eachCell((cell, colNum) => {
  //         if (colNum === 1) {
  //           cell.alignment = { horizontal: 'left', vertical: 'middle' };
  //         } else {
  //           const val = cell.value;
  //           if (val == null || val === '-') {
  //             cell.value = '-';
  //           } else if (typeof val === 'number') {
  //             if (val < 0)
  //               cell.value = `($${Math.abs(val).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })})`;
  //             else
  //               cell.value = `$${val.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
  //           }
  //           cell.alignment = { horizontal: 'right', vertical: 'middle' };
  //         }
  //       });
  //     });
  //   }
  // });



  //   worksheet.columns.forEach((col: any, i: number) => {
  //     col.width = i === 0 ? 30 : 15;
  //   });


  //   this.shared.exportToExcel(workbook, `Financial Statement_${currentMonth}`);
  // }

  formatExcelValue(value: any, valueType: string) {
    if (value === null || value === '' || value === undefined || value === 0 || value === '0') {
      return '-';
    }
    switch (valueType) {
      case '#':
        return value;
      case '$':
        return value;
      case '%':
        return value / 100;
      default:
        return value;
    }
  }
  exportToExcel(): void {
    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('Financial Statement');

    const formattedMonth = this.shared.datePipe.transform(
      this.selectedDate,
      'MMMM yyyy'
    );

    // ================= STORE VALUE FOR EXCEL =================
    let storeValue = '';

    const allStores: any[] = Array.isArray(this.stores) ? this.stores : [];
    const selectedStoreIds: string[] = Array.isArray(this.storeIds)
      ? this.storeIds.map((id: any) => id.toString())
      : [];

    // ✅ No selection OR all selected → show ALL store names
    if (!selectedStoreIds.length || selectedStoreIds.length === allStores.length) {
      storeValue = allStores
        .map((s: any) => s.storename?.trim())
        .filter(Boolean)
        .join(', ');
    }
    // ✅ Partial selection → show ONLY selected store names
    else {
      storeValue = allStores
        .filter((s: any) => selectedStoreIds.includes(s.ID.toString()))
        .map((s: any) => s.storename?.trim())
        .filter(Boolean)
        .join(', ');
    }

    // ================= FILTERS =================
    const filters = [
      { name: 'Store:', values: storeValue },
      { name: 'Month:', values: formattedMonth },
    ];

    const currentMonth = formattedMonth;

    // ================= TITLE =================
    worksheet.mergeCells('A1:G1');
    const title = worksheet.getCell('A1');
    title.value = 'Financial Statement';
    title.alignment = { horizontal: 'center', vertical: 'middle' };
    title.font = { bold: true, size: 14 };
    worksheet.getRow(1).height = 25;

    // ================= FILTER ROWS =================
    let startIndex = 2;
    filters.forEach(filter => {
      const labelCell = worksheet.getCell(`A${startIndex}`);
      const valueCell = worksheet.getCell(`B${startIndex}`);

      labelCell.value = filter.name;
      labelCell.font = { bold: true };
      labelCell.alignment = { horizontal: 'left' };

      worksheet.mergeCells(`B${startIndex}:C${startIndex}`);
      valueCell.value = filter.values;
      valueCell.alignment = { wrapText: true };

      startIndex++;
    });

    worksheet.addRow([]);

    // ================= HEADER =================
    const dynamicMonths = this.lastFourMonthsArray.map(m =>
      this.shared.datePipe.transform(m, 'MMMM yyyy')
    );

    const headerRow = [
      'Department',
      'Actual',
      'Budget',
      'Variance',
      ...dynamicMonths
    ];

    const headerRow1 = worksheet.addRow(headerRow);

    headerRow1.eachCell({ includeEmpty: true }, cell => {
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.alignment = { horizontal: 'center' };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF2F5597' }
      };
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      };
    });

    // ================= DATA =================
    this.FSData.forEach((data: any) => {
      const mainValues = [
        data.DisplayName || '',
        data.Actual,
        data.Budget,
        data.Variance,
        data.LM1,
        data.LM2,
        data.LM3,
      ];

      const mainRow = worksheet.addRow(mainValues);

      if (data.IsHeader === 'T' || data.IsHeader === 'Y') {
        const color = data.IsHeader === 'T' ? 'FFDDEBF7' : 'FFBFD4F2';
        mainRow.font = { bold: true };
        mainRow.eachCell(cell => {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: color } };
          cell.font = { bold: true };
          cell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
          };
        });
      }

      mainRow.eachCell({ includeEmpty: true }, (cell, colNum) => {
        if (colNum === 1) {
          cell.alignment = { horizontal: 'left' };
          return;
        }

        const formatted = this.formatExcelValue(cell.value, data.ValueType);

        if (formatted !== '-') {
          cell.value = formatted;
          if (data.ValueType === '$') cell.numFmt = '$#,##0.00';
          else if (data.ValueType === '#') cell.numFmt = '#,##0';
          else if (data.ValueType === '%') cell.numFmt = '0.00%';
        }

        cell.alignment = { horizontal: 'right' };
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      });

      // ================= SUB DATA =================
      if (data.SubData?.length) {
        data.SubData.forEach((sub: any) => {
          const subValues = [
            `   ${sub['SUBTYPE DETAIL'] || ''}`,
            sub.Actual,
            sub.Budget,
            sub.Variance,
            sub.LM1,
            sub.LM2,
            sub.LM3,
          ];

          const subRow = worksheet.addRow(subValues);
          subRow.font = { size: 10 };

          subRow.eachCell({ includeEmpty: true }, (cell, colNum) => {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'FFE8EEF7' }
            };

            cell.border = {
              top: { style: 'thin' },
              left: { style: 'thin' },
              bottom: { style: 'thin' },
              right: { style: 'thin' }
            };

            if (colNum === 1) {
              cell.alignment = { horizontal: 'left' };
              return;
            }

            const formatted = this.formatExcelValue(cell.value, sub.ValueType);

            if (formatted !== '-') {
              cell.value = formatted;
              if (sub.ValueType === '$') cell.numFmt = '$#,##0.00';
              else if (sub.ValueType === '#') cell.numFmt = '#,##0';
              else if (sub.ValueType === '%') cell.numFmt = '0.00%';
            }

            cell.alignment = { horizontal: 'right' };
          });
        });
      }
    });

    // ================= COLUMN STYLES =================
    worksheet.getColumn(1).alignment = { horizontal: 'left' };
    worksheet.columns.forEach((col, i) => {
      col.width = i === 0 ? 30 : 18;
    });

    // ================= EXPORT =================
    this.shared.exportToExcel(
      workbook,
      `Financial Statement_${currentMonth}`
    );
  }





}
