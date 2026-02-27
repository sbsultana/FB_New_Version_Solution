
import { Component, HostListener, OnInit, signal } from '@angular/core';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { Api } from '../../../../Core/Providers/Api/api';
import * as XLSX from 'xlsx';
import { Title } from '@angular/platform-browser';
import { common } from '../../../../common';

@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [SharedModule],
  templateUrl: './dashboard.html',
  styleUrls: ['./dashboard.scss']
})
export class Dashboard implements OnInit {
  activePopover: number | null = null;
  displayYear: number = new Date().getFullYear();
  reportData: any[] = [];
  detailsData: any[] = [];
  storeButtonLabel = 'All';
  isPopoverOpen: boolean = false;
  stores: any[] = [];
  isMonthDropdownOpen: boolean = false;
  selectedMonthLabel: string = '';
  months: string[] = [];
  selectedMonthIndex: number = new Date().getMonth();
  selectedMonth: Date | null = null;
  deptOrder: { [key: string]: number } = { New: 1, Used: 2 };
  expandedDept: string | null = null;
  noData: boolean = false;
  noDetails: boolean = false;
  storeIds: number[] = [];
  selectedDept: any = '';
  reportTotal: any = null;
  searchText: string = '';
  showModal = false;
  searchQuery: string = '';
  accountDetails: any[] = [];
  selectedRow: any = null;
  selectedStoreName: string = '';
  selectedAccountNumber: string = '';
  showDetailPage: boolean = false;
  currentPage: number = 1;
  pageSize: number = 20;
  Math = Math;
  totalRecords: number = 0;
  totalPagesCount: number = 0;
  pagedData: any[] = [];
  noAccountDetails: boolean = false;
  details = signal<any[]>([]);
  originalReportData: any[] = [];
  filteredDetails: any[] = [];
  selectedDepartments: string[] = ['NEW', 'USED'];
  selectedReportTotal: string = 'BOTTOM';

  constructor(
    public shared: Sharedservice,
    private api: Api, private title: Title,
     private comm: common,   public apiSrvc: Api,
  ) { 
    this.title.setTitle(this.comm.titleName + '-Enterprise Tracking');
    const data = {
      title: 'Variable Gross GL',
      
    };
    this.apiSrvc.SetHeaderData({ obj: data });
  }

  ngOnInit(): void {
    this.months = Array.from({ length: 12 }, (_, i) =>
      new Date(0, i).toLocaleString('en', { month: 'long' })
    );
    this.selectedMonthLabel = `${this.months[this.selectedMonthIndex]}`;
    this.getStoresList();
    this.getVariableGrossGLReport();
    this.filteredDetails = [...this.detailsData];
    // this.getVariableGrossGLAccountDetails();
  }

  // ---------- PAGINATION ----------
  updatePagedData() {
    const start = (this.currentPage - 1) * this.pageSize;
    const end = start + this.pageSize;
    this.pagedData = this.reportData.slice(start, end);
  }

  previousPage() {
    if (this.currentPage > 1) {
      this.currentPage--;
      this.updatePagedData();
    }
  }
  getTotalBalance(): number {
    if (!this.details || typeof this.details !== 'function') return 0;

    const list = this.details();
    if (!Array.isArray(list) || list.length === 0) return 0;

    // Sum all valid numeric balances
    const total = list.reduce((sum: number, item: any) => {
      const balance = parseFloat(item?.balance);
      return sum + (isNaN(balance) ? 0 : balance);
    }, 0);

    return total;
  }

  getStoresList(): Promise<void> {
    return new Promise((resolve) => {
      this.shared.spinner.show();
      this.api.postmethod('Western/GetStoresList', { userid: 1 }).subscribe({
        next: (res: any) => {
          this.shared.spinner.hide();
          const storeData = res?.response || res || [];
          this.stores = storeData.map((s: any) => ({
            ID: s.ID,
            storename: s.storename || s.StoreName
          }));
          this.storeIds = this.stores.map((s) => s.ID);
          this.updateStoreSelection();
          resolve();
        },
        error: () => {
          this.shared.spinner.hide();
          this.stores = [];
          this.storeIds = [];
          resolve();
        }
      });
    });
  }

  allstores(state: 'Y' | 'N') {
    this.storeIds = state === 'Y' ? this.stores.map(s => s.ID) : [];
    this.updateStoreSelection();
    this.getVariableGrossGLReport();
  }

  individualStores(store: any) {
    const idx = this.storeIds.indexOf(store.ID);
    idx >= 0 ? this.storeIds.splice(idx, 1) : this.storeIds.push(store.ID);
    this.updateStoreSelection();
    this.getVariableGrossGLReport();
  }

  onStoreToggle(storeId: number) {
    const idx = this.storeIds.indexOf(storeId);
    if (idx >= 0) this.storeIds.splice(idx, 1);
    else this.storeIds.push(storeId);
    this.updateStoreSelection();
    this.getVariableGrossGLReport();
  }

  updateStoreSelection() {
    if (this.storeIds.length === this.stores.length) {
      this.storeButtonLabel = 'All';
    } else if (this.storeIds.length === 1) {
      const selected = this.stores.find(s => s.ID === this.storeIds[0]);
      this.storeButtonLabel = selected?.storename || '';
    } else {
      this.storeButtonLabel = this.storeIds.length.toString();
    }
  }

  formatAPIDate(date: any): string {
    const d = new Date(date);
    const year = d.getFullYear();
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  }

  changeYear(amount: number) {
    // this.displayYear += amount;
    this.selectedMonthLabel = `${this.months[this.selectedMonthIndex]}`;
  }

  isFutureMonth(index: number): boolean {
    const current = new Date();
    const selected = new Date(this.displayYear, index, 1);
    return selected > current;
  }

  selectMonth(index: number) {
    this.selectedMonthIndex = index;
    this.tempMonthLabel = `${this.months[index]}`;
    this.isMonthDropdownOpen = false;

    // Store date temporarily for Apply
    this.selectedDate = new Date(this.displayYear, index, 1);
    this.getVariableGrossGLReport();
  }

  tempMonthLabel: string = '';
  selectedDate: Date | null = null;
  loading=false;

  async getVariableGrossGLReport() {
    this.shared.spinner.show();
    this.loading=true;
    if (!this.stores || this.stores.length === 0) {
      await this.getStoresList();
    }

    const date = this.formatDate(new Date(this.selectedDate || new Date()));

    let storeIds = '';
    if (this.storeIds && this.storeIds.length > 0) {
      storeIds = this.storeIds.join(',');
    } else if (this.stores && this.stores.length > 0) {
      storeIds = this.stores.map((s: any) => s.ID).join(',');
    }

    const payload = {
      DATE: date,
      AS_IDS: storeIds,
      DEALTYPE: ''
    };

    this.api.postmethod('WesternAuto/GetSalesGrossGLSummaryReport', payload).subscribe({
      next: (res: any) => {
        this.shared.spinner.hide();

        const raw = Array.isArray(res?.response) ? res.response : [];
        if (!raw.length) {
          this.noData = true;
          this.reportData = [];
          this.reportTotal = null;
          return;
        }

        // ✅ Separate out "Report Total" and "Other Stores"
        const reportTotals = raw.filter((x: any) => x.STORE === 'Report Total');
        const storeRows = raw.filter((x: any) => x.STORE !== 'Report Total');

        // ✅ Group each store and its departments
        const groupedData = storeRows.reduce((acc: any[], item: any) => {
          let group = acc.find(g => g.STORE === item.STORE);
          if (!group) {
            group = { STORE: item.STORE, children: [] };
            acc.push(group);
          }

          if (item.DEPT === null) {
            Object.assign(group, item);
          } else {
            group.children.push(item);
          }

          return acc;
        }, []);

        // ✅ Add "Report Total" as a visible group with its children
        if (reportTotals.length) {
          const reportParent = reportTotals.find((x: any) => x.DEPT === null);
          const reportChildren = reportTotals.filter((x: any) => x.DEPT !== null);

          if (reportParent) {
            groupedData.push({
              ...reportParent,
              STORE: 'Report Total',
              children: reportChildren,
              alwaysVisible: true // mark as auto-visible
            });
          }
        }

        // ✅ Assign to table data
        this.reportData = groupedData;
        this.originalReportData = JSON.parse(JSON.stringify(groupedData));
        this.reportTotal = null;
        this.noData = false;
        
      },
      error: () => {
        this.shared.spinner.hide();
        this.noData = true;
      }
    });
  }

onPageSizeChange() {
  this.currentPage = 1;
  this. getVariableGrossGLDetails(this.selectedDept); // your API or data fetching method
}
  formatDate(date: Date): string {
    const d = new Date(date);
    const month = `${d.getMonth() + 1}`.padStart(2, '0');
    const year = d.getFullYear();
    return `${year}-${month}-01`;
  }

  // ---------------- DROPDOWNS ----------------
  toggleDepartmentDropdown() {
    this.isDepartmentDropdownOpen = !this.isDepartmentDropdownOpen;
    this.isReportDropdownOpen = false;
    this.isMonthDropdownOpen = false;
  }

  toggleReportDropdown() {
    this.isReportDropdownOpen = !this.isReportDropdownOpen;
    this.isDepartmentDropdownOpen = false;
    this.isMonthDropdownOpen = false;
  }

  toggleMonthDropdown() {
    this.isMonthDropdownOpen = !this.isMonthDropdownOpen;
    this.isDepartmentDropdownOpen = false;
    this.isReportDropdownOpen = false;
  }

  togglePopover(index: number) {
    this.activePopover = this.activePopover === index ? null : index;
    this.isDepartmentDropdownOpen = false;
    this.isReportDropdownOpen = false;
    this.isMonthDropdownOpen = false;
  }

  // ---------------- SELECTORS ----------------
  toggleDepartment(dept: string) {
    const index = this.selectedDepartments.indexOf(dept);
    if (index > -1) this.selectedDepartments.splice(index, 1);
    else this.selectedDepartments.push(dept);
  }

  selectReportTotal(position: 'TOP' | 'BOTTOM') {
    this.selectedReportTotal = position;
    this.isReportDropdownOpen = false;
  }

  // ✅ Close all dropdowns when clicking outside
  @HostListener('document:click', ['$event'])
  onClickOutside(event: Event) {
    const target = event.target as HTMLElement;
    if (!target.closest('.dropdown')) {
      this.isMonthDropdownOpen = false;
      this.isDepartmentDropdownOpen = false;
      this.isReportDropdownOpen = false;
      this.activePopover = null;
    }
  }

  openModal(child: any) {
    if (!child?.DEPT) return;

    this.selectedDept = child.DEPT;
    this.getVariableGrossGLDetails(this.selectedDept);
    this.showModal = true; // show modal
  }
  isDropdownOpen = false;
  totalBalance: number = 0;
  getVariableGrossGLDetails(dept: string) {
  this.noDetails = false;
  this.totalRecords = 0;
  this.totalBalance = 0; // reset total

  const payload = {
    AS_IDS: this.storeIds?.length ? this.storeIds.join(',') : '',
    DATE: this.selectedDate || new Date().toISOString().split('T')[0],
    VAR1: dept,
    Accountnumber: ''
  };

  console.log('Payload:', payload);

  this.api.postmethod('WesternAuto/GetSalesGrossGLDetails', payload).subscribe({
    next: (res) => {
      const response = res?.response || [];

      // ✅ Map API response
      const mapped = response.map((item: any) => ({
        store: item.Store,
        accountNumber: item.AccountNumber,
        accountDescription: item['Account Description'],
        balance: Number(item.Balance) || 0 // ensure numeric
      }));

      // ✅ Assign data
      this.tableData = mapped;
      this.details.set(mapped);

      // ✅ Calculate total
      this.calculateTotal();

      // ✅ Reset pagination to first page and update
      this.currentPage = 1;
      this.updatePagination();

      console.log('Total Balance:', this.totalBalance);
    },
    error: () => {
      this.details.set([]);
      this.totalBalance = 0;
      this.totalRecords = 0;
      this.pagedDetails = [];
      this.currentPage = 1;
    }
  });
}

  tableData: any[] = [];

  calculateTotal() {
    if (this.tableData && this.tableData.length > 0) {
      this.totalBalance = this.tableData.reduce(
        (sum: any, item: { balance: any; }) => sum + (item.balance || 0),
        0
      );
    } else {
      this.totalBalance = 0;
    }
  }

  selectRow(row: any) {
    this.showDetailPage = !this.showDetailPage;
    if (this.selectedRow === row) {
      // Toggle off if the same row is clicked
      this.selectedRow = null;
      this.accountDetails = [];
    } else {
      this.selectedRow = row;
      this.selectedStoreName = row.Store || '';
      this.selectedAccountNumber = row.AccountNumber || '';
      this.accountDetails = []; // Set empty or fetch API data here
    }

    this.getVariableGrossGLAccountDetails(row)
  }

  positions = ['Top', 'Bottom'];
  selectedPosition: string = '';
  isDepartmentDropdownOpen = false;
  isReportDropdownOpen = false;
  departments: string[] = ['NEW', 'USED'];

  toggleDropdown(): void {
    this.isDropdownOpen = !this.isDropdownOpen;
  }


//  getVariableGrossGLAccountDetails(row: any): void {
//   // keep selectedRow as set by onAccountClick
//   this.noAccountDetails = false;
//   this.accountDetails = [];
//   this.showDetailPage = false;

//   const payload = {
//     AS_IDS: this.storeIds?.length ? this.storeIds.join(',') : '',
//     DATE: this.selectedDate || new Date().toISOString().split('T')[0],
//     Accountnumber: row?.accountNumber || row?.AccountNumber || '',
//     VAR1: this.selectedDept || '',
//     VAR2: ''
//   };

//   console.log('Payload for account details:', payload);

//   this.shared.spinner.show();
//   this.api.postmethod('WesternAuto/GetSalesGrossGLDetails', payload).subscribe({
//     next: (res: any) => {
//       this.shared.spinner.hide();

//       const resp = Array.isArray(res?.response) ? res.response : [];

//       this.accountDetails = resp.map((item: any) => ({
//         Control: item.control ?? item.Control ?? '-',
//         AccountingDate: item.Date ? new Date(item.Date).toISOString().split('T')[0] : '-',
//         DetailDescription: item['Account Description'] || item['AccountDescription'] || '-',
//         PostingAmount: Number(item.Balance) || 0,
//         AsDealerName: item.As_dealername || item.AsDealerName || '-'
//       }));

//       this.calculateSubTotal();
//       this.noAccountDetails = this.accountDetails.length === 0;
//       this.showDetailPage = true;
//     },
//     error: (err) => {
//       this.shared.spinner.hide();
//       console.error('Error loading account details:', err);
//       this.accountDetails = [];
//       this.noAccountDetails = true;
//       this.showDetailPage = false;
//     }
//   });
// }

  // ---------- UI LOGIC ----------
  toggleAccountDetails(row: any) {
    if (this.expandedRow === row) {
      this.expandedRow = null;
      this.showDetailPage = false;
    } else {
      this.expandedRow = row;
      this.selectedStoreName = row.StoreName;
      this.selectedAccountNumber = row.AccountNumber;
      this.showDetailPage = true;

    }
  }


// ✅ Slice the data for current page
setPagedDetails() {
  const data = this.details(); // or your main table array
  const startIndex = (this.currentPage - 1) * this.pageSize;
  const endIndex = startIndex + this.pageSize;
  this.pagedDetails = data.slice(startIndex, endIndex);
}

// ✅ Navigation functions
nextPage() {
  if (this.currentPage < this.totalPages()) {
    this.currentPage++;
    this.setPagedDetails();
  }
}

prevPage() {
  if (this.currentPage > 1) {
    this.currentPage--;
    this.setPagedDetails();
  }
}

goToPage(page: number) {
  if (page >= 1 && page <= this.totalPages()) {
    this.currentPage = page;
    this.setPagedDetails();
  }
}

totalPages() {
  return Math.ceil(this.totalRecords / this.pageSize) || 1;
}

pagedDetails: any[] = [];
  scrollToRow(index: number) {
    const el = document.getElementById('row-' + index);
    if (el) el.scrollIntoView({ behavior: 'smooth', block: 'center' });
  }

  // ---------- PAGINATION ----------
  paginatedDetails() {
    const start = (this.currentPage - 1) * this.pageSize;
    return this.detailsData.slice(start, start + this.pageSize);
  }

  // ---------- TOTALS ----------
  getSubTotal(list: any[]): number {
    return list.reduce((sum, item) => sum + (item.PostingAmount || 0), 0);
  }

  // ---------- MODAL ----------
  closeModal() {
    this.showModal = false;
    this.detailsData = [];
    this.accountDetails = [];
  }

  applyFilters() {
    // ✅ Apply selected month officially
    if (this.tempMonthLabel) {
      this.selectedMonthLabel = this.tempMonthLabel;
    }

    // 🔹 (Rest of your filter logic remains unchanged)
    if (!this.originalReportData || this.originalReportData.length === 0) {
      this.noData = true;
      return;
    }

    let filteredData = JSON.parse(JSON.stringify(this.originalReportData));

    if (this.selectedDepartments.length > 0) {
      filteredData = filteredData
        .map((store: any) => ({
          ...store,
          children: store.children.filter((child: any) =>
            this.selectedDepartments.includes(child.DEPT?.toUpperCase())
          )
        }))
        .filter((store: any) => store.children.length > 0 || store.alwaysVisible);
    }

    const reportTotalIndex = filteredData.findIndex(
      (g: any) => g.STORE === 'Report Total'
    );
    if (reportTotalIndex !== -1) {
      const [reportTotal] = filteredData.splice(reportTotalIndex, 1);
      if (this.selectedReportTotal === 'TOP') filteredData.unshift(reportTotal);
      else filteredData.push(reportTotal);
    }

    this.reportData = filteredData;
    this.noData = this.reportData.length === 0;
  }
  subTotal: number = 0;

  calculateSubTotal() {
    this.subTotal = this.accountDetails.reduce(
      (sum, detail) => sum + (detail.PostingAmount || 0),
      0
    );
  }

onSearchChange(query: string) {
  const q = (query || '').toString().trim().toLowerCase();

  // prefer using the signal () if it exists, otherwise fallback to tableData
  const source: any[] = Array.isArray(this.details?.()) ? this.details() : (this.tableData || []);

  if (!q) {
    // reset to full source
    this.filteredDetails = [...source];
  } else {
    this.filteredDetails = source.filter((item: any) => {
      const store = (item.store || item.Store || '').toString().toLowerCase();
      const acct = (item.accountNumber || item.AccountNumber || '').toString().toLowerCase();
      const desc = (item.accountDescription || item['Account Description'] || '').toString().toLowerCase();
      return store.includes(q) || acct.includes(q) || desc.includes(q);
    });
  }

  // reset pagination for filtered data
  this.currentPage = 1;
  this.updatePagination();
}
updatePagination() {
  // source preference: filteredDetails (if set) -> details() signal -> tableData
  let source: any[] =
    (this.filteredDetails && this.filteredDetails.length > 0) ? this.filteredDetails
    : (Array.isArray(this.details?.()) ? this.details() : (this.tableData || []));

  // When filteredDetails exists but is empty (search returned no results), source should be filteredDetails
  if (this.filteredDetails && this.filteredDetails.length === 0 && this.searchText) {
    source = this.filteredDetails;
  }

  this.totalRecords = source.length;
  // compute pagedDetails slice
  const startIndex = (this.currentPage - 1) * this.pageSize;
  const endIndex = startIndex + this.pageSize;
  this.pagedDetails = source.slice(startIndex, endIndex);
}
expandedRow: any = null;

onAccountClick(row: any): void {
  // If the same row is clicked again, collapse it
  if (this.expandedRow === row) {
    this.expandedRow = null;
    this.showDetailPage = false;
    this.accountDetails = [];
    return;
  }

  // Otherwise, expand a new row
  this.expandedRow = row;
  this.showDetailPage = false; // Hide detail until data loads
  this.getVariableGrossGLAccountDetails(row);
}


getVariableGrossGLAccountDetails(row: any): void {
  this.noAccountDetails = false;
  this.accountDetails = [];

  const payload = {
    AS_IDS: this.storeIds?.length ? this.storeIds.join(',') : '',
    DATE: this.selectedDate || new Date().toISOString().split('T')[0],
    Accountnumber: row?.accountNumber || row?.AccountNumber || '',
    VAR1: this.selectedDept || '',
    VAR2: ''
  };

  console.log('Payload for account details:', payload);

  this.shared.spinner.show();

  this.api.postmethod('WesternAuto/GetSalesGrossGLDetails', payload).subscribe({
    next: (res: any) => {
      this.shared.spinner.hide();

      const resp = Array.isArray(res?.response) ? res.response : [];

      this.accountDetails = resp.map((item: any) => ({
        Control: item.control ?? item.Control ?? '-',
        AccountingDate: item.Date
          ? new Date(item.Date).toISOString().split('T')[0]
          : '-',
        DetailDescription: item['Account Description'] || item['AccountDescription'] || '-',
        PostingAmount: Number(item.Balance) || 0,
        AsDealerName: item.As_dealername || item.AsDealerName || '-'
      }));

      this.calculateSubTotal();
      this.noAccountDetails = this.accountDetails.length === 0;
      this.showDetailPage = true;
    },
    error: (err) => {
      this.shared.spinner.hide();
      console.error('Error loading account details:', err);
      this.accountDetails = [];
      this.noAccountDetails = true;
      this.showDetailPage = false;
    }
  });
}



downloadDetails() {
  if (!this.tableData || this.tableData.length === 0) {
    alert('No data to download');
    return;
  }

  const wsData = [
    ['STORE NAME', 'ACCOUNT NUMBER', 'ACCOUNT DESCRIPTION', 'BALANCE ($)'],
    ...this.tableData.map(row => [
      row.store,
      row.accountNumber,
      row.accountDescription,
      row.balance
    ])
  ];

  const ws: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet(wsData);

  // Auto width
  ws['!cols'] = [
    { wch: 25 },
    { wch: 20 },
    { wch: 40 },
    { wch: 15 }
  ];

  const wb: XLSX.WorkBook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Details');
  XLSX.writeFile(wb, 'Details_Report.xlsx');
}




  // ---------- HELPER ----------
  getMonthYearLabel(dateString: string): string {
    if (!dateString) return '';
    const date = new Date(dateString);
    const month = date.toLocaleString('default', { month: 'long' }); // e.g. 'November'
    const year = date.getFullYear();
    return `${month}/${year}`; // Matches 'NOVEMBER/2025' format
  }


  toggleDetails(dept: string) {
    this.expandedDept = this.expandedDept === dept ? null : dept;
    this.getVariableGrossGLDetails(dept);
  }
}
