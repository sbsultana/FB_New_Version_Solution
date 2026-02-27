import { ChangeDetectorRef, Component, ElementRef, HostListener, Inject, PLATFORM_ID, ViewChild } from '@angular/core';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { ToastContainer } from '../../../../Layout/toast-container/toast-container';
import { NgxSpinnerModule, NgxSpinnerService } from 'ngx-spinner';
import { BsDatepickerConfig, BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
import { Api } from '../../../../Core/Providers/Api/api';
import { HttpClient } from '@angular/common/http';
import { NgbModal } from '@ng-bootstrap/ng-bootstrap';
import { common } from '../../../../common';
import { FormBuilder } from '@angular/forms';
import { Router } from '@angular/router';

@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [SharedModule, ToastContainer, NgxSpinnerModule, BsDatepickerModule],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss'
})
export class Dashboard {
  @ViewChild('chartContainer3') chartContainer3!: ElementRef<HTMLDivElement>;
  @ViewChild('chartContainer2') chartContainer2!: ElementRef<HTMLDivElement>;
  // selectedYear: Date = new Date();

  // isEditMode = false;

  // Original data
  AbingtonData = [
    { variable: 'Wholesale', months: [1200, 1500, 1800, 2000, 1700, 2200, 2400, 2300, 2100, 2500, 2600, 2700] },
    { variable: 'Inter-Company', months: [800, 750, 900, 1000, 950, 1100, 1050, 980, 1020, 1200, 1150, 1250] },
  ];

  PDCData = [
    { variable: 'Wholesale', months: [1200, 1500, 1800, 2000, 1700, 2200, 2400, 2300, 2100, 2500, 2600, 2700] },
    { variable: 'Dealer', months: [800, 750, 900, 1000, 950, 1100, 1050, 980, 1020, 1200, 1150, 1250] },
    { variable: 'Fleet', months: [400, 750, 900, 1000, 750, 1100, 1350, 1320, 1080, 1300, 1450, 1450] },
    { variable: 'Bulk', months: [2000, 2200, 2400, 2500, 2600, 2800, 3000, 3200, 3100, 3300, 3400, 3500] },
    { variable: 'Aftermarket Powertrain', months: [5000, 0, 2000, 0, 0, 4500, 0, 0, 3000, 0, 1000, 0] },
    { variable: 'Retail / Internet', months: [400, 750, 900, 1000, 750, 1100, 1350, 1320, 1080, 1300, 1450, 1450] },
    { variable: 'Inter-Company', months: [2000, 2200, 2400, 2500, 2600, 2800, 3000, 3200, 3100, 3300, 3400, 3500] },
  ];
  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent) {
    const clickedInside = (event.target as HTMLElement).closest('.dropdown-toggle, .reportstores-card , .timeframe');
    if (!clickedInside) {
      this.activePopover = -1;
    }
  }
  tableData: { variable: string; months: number[] }[] = [];
  // backupData: { variable: string; months: number[] }[] = [];
  constructor(
    public shared: Sharedservice,
    @Inject(PLATFORM_ID) private platformId: Object,
    private toastService: ToastService,
    private spinner: NgxSpinnerService,
    public apiSrvc: Api,
    private http: HttpClient,
    public ngbModal: NgbModal,
    private comm: common,
    private fb: FormBuilder,
    private router: Router,
    private cd: ChangeDetectorRef,
  ) {
    this.shared.setTitle(this.shared.common.titleName + '- Budget Forecast');
    let obj = {
      title: 'Budget Forecast',
    }
    this.apiSrvc.SetHeaderData({ obj });
    this.tableData = this.PDCData;
    this.bsConfig = {
      dateInputFormat: 'YYYY',
      minMode: 'year',
      maxDate: this.maxDate
    };
    this.getBudgetForecastData(this.selectedFilter, this.selectedYear);
  }
  PDCForecastData: any[] = [];
  NoPDCForecastdata = false;
  selectedstorevalues: number = 60;
  variableType: string = '';
  isEditMode = false;
  backupData: any;

  filters: string[] = ['store', 'expenses'];
  selectedFilter: string = 'store';

  activePopover: number | null = null;
  selectedYear!: number;

  bsConfig!: Partial<BsDatepickerConfig>;
  maxDate: Date = new Date();
  bsValue: Date = new Date();

  togglePopover(id: number) {
    this.activePopover = this.activePopover === id ? null : id;
  }
  selectFilter(f: string) {
    this.selectedFilter = f;
    this.activePopover = null;
  }
  onYearSelected(event: Date) {
    if (event) {
      this.selectedYear = event.getFullYear();
    }
  }
  applyFilterAndYear() {
    if (!this.selectedYear) {
      alert("Please select a year!");
      return;
    }
    this.getBudgetForecastData(this.selectedFilter, this.selectedYear);
  }
  cancelEdit() {
    this.isEditMode = false;

    // Restore from backup if available
    if (this.backupData && this.backupData.length) {
      this.PDCForecastData = JSON.parse(JSON.stringify(this.backupData));
    } else {
      // If no backup, just clear monthly fields
      const monthKeys = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
      this.PDCForecastData.forEach(row => {
        monthKeys.forEach(key => {
          row[key] = '';
        });
      });
    }
  }



  // getRowTotal(row: { variable: string; months: number[] }): number {
  //   return row.months.reduce((sum, val) => sum + val, 0);
  // }

  // getGrandTotal(): number {
  //   return this.tableData.reduce((acc, row) => acc + this.getRowTotal(row), 0);
  // }
  // getColumnTotal(index: number): number {
  //   return this.tableData.reduce((sum, row) => sum + (row.months[index] || 0), 0);
  // }

  // Add these helper functions in your component (below getBudgetForecastData)

  getRowTotal(row: any): number {
    const months = [
      row.Jan, row.Feb, row.Mar, row.Apr, row.May, row.Jun,
      row.Jul, row.Aug, row.Sep, row.Oct, row.Nov, row.Dec
    ].map((m: any) => Number(m) || 0);
    return months.reduce((a: number, b: number) => a + b, 0);
  }

  getColumnTotal(index: number): number {
    const monthKeys = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const key = monthKeys[index];
    return this.PDCForecastData.reduce((sum, row) => sum + (Number(row[key]) || 0), 0);
  }

  getGrandTotal(): number {
    return this.PDCForecastData.reduce((sum, row) => sum + (row.sum || 0), 0);
  }

  selectedDt: number = 2025;

  getBudgetForecastData(type: string, year: number): void {

    this.variableType = type;

    const obj = {
      IV_StoreID: 2,
      IV_V_Type: type,
      IV_Year: year?.toString() || new Date().getFullYear().toString()
    };

    this.spinner.show();
    this.apiSrvc.postmethod('BudgetValues/GetBudgetValuesExpense', obj)
      .subscribe({
        next: (res: any) => {
          this.spinner.hide();
          if (res.status === 200 && Array.isArray(res.response)) {
            this.PDCForecastData = [...res.response].sort((a: any, b: any) => {
              const catA = a.CategorySequence ?? 0;
              const catB = b.CategorySequence ?? 0;
              const seqA = a.SequenceNo ?? 0;
              const seqB = b.SequenceNo ?? 0;
              return catA === catB ? seqA - seqB : catA - catB;
            });
            this.cd.detectChanges();
            const userId = JSON.parse(localStorage.getItem('UserDetails') || '{}')?.userid;

            this.PDCForecastData.forEach(ele => {
              if (!ele.IV_Variable_ID || ele.IV_Variable_ID === 'null') {
                ele.IV_Variable_ID = ele.ITV_Id;
                ele.IV_StoreID = this.selectedstorevalues;
                ele.IV_Seq = ele.SequenceNo ?? 0;
                ele.IV_Formula_Status = ele.ITV_Formula_Status;
                ele.IV_User_ID = userId;
                ele.IV_Year = year.toString();
              }

              const monthKeys = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

              // Keep empty if no value, otherwise convert to number
              monthKeys.forEach(key => {
                ele[key] = ele[key] != null && ele[key] !== '' ? Number(ele[key]) : null;
              });

              // Calculate sum treating empty as 0
              const monthValues = monthKeys.map(key => Number(ele[key]) || null);
              ele.sum = monthValues.reduce((a: any, b: any) => a + b, null);
            });
          } else {
            this.PDCForecastData = [];
          }
          this.NoPDCForecastdata = this.PDCForecastData.length === 0;
        },
        error: err => {
          console.error('API Error:', err);
          this.spinner.hide();
          this.NoPDCForecastdata = true;
        }
      });
  }

  EditSave() {
    if (!this.isEditMode) {
      this.backupData = JSON.parse(JSON.stringify(this.PDCForecastData));
      this.isEditMode = true;
    } else {
      const storeId = 2;
      const userId = JSON.parse(localStorage.getItem('UserDetails') || '{}')?.userid;

      const payload = this.PDCForecastData.map(ele => {
        const newRow = { ...ele };

        if (!newRow.IV_Variable_ID || newRow.IV_Variable_ID === 'null') {
          newRow.IV_Variable_ID = newRow.ITV_Id;
          newRow.IV_StoreID = storeId;
          newRow.IV_Seq = newRow.SequenceNo ?? 0;
          newRow.IV_Formula_Status = newRow.ITV_Formula_Status;
          newRow.IV_User_ID = userId;
          newRow.IV_Year = this.selectedYear.toString();
        } else {
          newRow.IV_StoreID = storeId;
        }
        const monthKeys = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
        monthKeys.forEach(key => {
          newRow[key] = newRow[key] != null && newRow[key] !== '' ? newRow[key] : null;
        });
        const monthValues = monthKeys.map(key => Number(newRow[key]) || null);
        newRow.sum = monthValues.reduce((a: any, b: any) => a + b, null);

        return newRow;
      });


      this.spinner.show();
      this.apiSrvc.postmethod('BudgetValues/AddBudgetValues', payload)
        .subscribe({
          next: res => {
            this.spinner.hide();
            if (res.status === 200) {
              this.toastService.show('Data saved successfully!', 'success', 'Success');
              this.onYearSelected(new Date(this.selectedYear, 0, 1));
              this.isEditMode = false;
            } else {
              this.toastService.show('Invalid inputs!', 'danger', 'Error');
            }
          },
          error: err => {
            console.error('API error: ', err);
            this.toastService.show('Something went wrong!', 'danger', 'Error');
            this.spinner.hide();
          }
        });
    }
  }
  monthsShort = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  onPaste(event: ClipboardEvent, rowIndex: number, startColIndex: number): void {
    event.preventDefault();

    const clipboardData = event.clipboardData || (window as any).clipboardData;
    const pastedText = clipboardData.getData('text');

    const rows = pastedText.split('\n').filter((r: any) => r.trim() !== '');
    const values = rows[0].split('\t'); // Only first row for now

    if (!this.PDCForecastData[rowIndex]) return;

    for (let i = 0; i < values.length; i++) {
      const colIndex = startColIndex + i;
      if (colIndex >= this.monthsShort.length) break;

      const monthKey = this.monthsShort[colIndex];
      this.PDCForecastData[rowIndex][monthKey] = values[i].trim();

      // Optional: call input handler if you have one
      // this.onKeydownEvent(new InputEvent('input'), this.gridObj[rowIndex], rowIndex, monthKey, colIndex);
    }
  }

  BacktoPDCDashboard() {
    this.router.navigate(['']);
  }


}
