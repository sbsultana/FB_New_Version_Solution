import { Component, HostListener } from '@angular/core';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { Api } from '../../../../Core/Providers/Api/api';
import * as XLSX from 'xlsx';
import { HttpClient } from '@angular/common/http';
import { Subscription } from 'rxjs';
import { Workbook } from 'exceljs';
import * as FileSaver from 'file-saver';
import { DatePipe, CommonModule } from '@angular/common';
@Component({
  selector: 'app-dashboard',
  imports: [SharedModule],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss'
})
export class Dashboard {
  isReportDropdownOpen = false;
  activePopover: number | null = null;
  storeButtonLabel = 'All';
  storeIds: any = '2'; // Default to store ID 2
  stores: any[] = [];
  selectedReportTotal: string = 'BOTTOM';
  storeData: any[] = [];
  showModal = false;
  selectedStore: any = null;
  selectedGroup: any = null;
  groupedStores: { [key: string]: any[] } = {};
  partsDetails: boolean = false;
  partsSourceList: any[] = []; // Stores the API response
  selectedStoreSource: any = [];
  isSourceDropdownOpen = false;
  partsDetailsList: any[] = [];
  originalStoreData: any[] = [];
  DateType: any = '30';
  excel!: Subscription;
  NoData: boolean = false;
 groupedParts: Record<string, any[]> = {};

  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent) {
    const clickedInside = (event.target as HTMLElement).closest('.filter-dropdowns, .dropdown-toggle, .reportstores-card , .timeframe');
    if (!clickedInside) {
      this.activePopover = -1;
    }
  }

  constructor(public shared: Sharedservice,
    private api: Api, private http: HttpClient,
    private datepipe: DatePipe,

  ) {
    this.shared.setTitle(this.shared.common.titleName + '-Parts Aging');
    if (localStorage.getItem('Fav') != 'Y') {
      const data = {
        title: 'Parts Aging',
        stores: this.storeIds,
        reportTotal: this.selectedReportTotal
      };
      this.shared.api.SetHeaderData({
        obj: data,
      });
    }
  }

  ngOnInit(): void {
    this.getStoresList();
    this.getPartsAgingReport();
    this.getPartsSource();
  }

  ngAfterViewInit(): void {
    this.excel = this.shared.api.getExportToExcelAllReports().subscribe((res: any) => {
      if (res && res.obj && res.obj.title == 'Parts Aging' && res.obj.state == true) {
        this.exportToExcel(); // merged export will create both sheets
      }
    });

  }


  getStoresList(): Promise<void> {
    return new Promise((resolve) => {
      this.shared.spinner.show();

      this.http.post('Western/GetStoresList', { userid: 1 }).subscribe({
        next: (res: any) => {
          this.shared.spinner.hide();
          const storeData = res?.response || res || [];

          // ✅ Group stores by sg_name
          const groupedStores: { [key: string]: any[] } = {};
          storeData.forEach((s: any) => {
            const groupName = s.sg_name || 'Others'; // fallback if sg_name missing
            if (!groupedStores[groupName]) {
              groupedStores[groupName] = [];
            }
            groupedStores[groupName].push({
              ID: s.ID,
              storename: s.storename || s.StoreName,
              sg_name: groupName
            });
          });

          // ✅ Flatten to full list
          this.stores = Object.values(groupedStores).flat();

          // ✅ Collect all store IDs dynamically
          this.storeIds = this.stores.map((s) => s.ID);

          // ✅ Optionally store grouped structure
          this.groupedStores = groupedStores; // for dropdown display or filtering

          // console.log('Grouped Stores:', this.groupedStores);
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


  toggleReportDropdown() {
    this.isReportDropdownOpen = !this.isReportDropdownOpen;
  }

  togglePopover(index: number) {
    this.activePopover = this.activePopover === index ? null : index;
    this.isReportDropdownOpen = false;
  }

  onStoreToggle(storeId: number) {
    const idx = this.storeIds.indexOf(storeId);
    if (idx >= 0) this.storeIds.splice(idx, 1);
    else this.storeIds.push(storeId);
    this.updateStoreSelection();
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

  allstores(state: 'Y' | 'N') {
    this.storeIds = state === 'Y' ? this.stores.map(s => s.ID) : [];
    this.updateStoreSelection();
  }

  individualStores(store: any) {
    const idx = this.storeIds.indexOf(store.ID);
    idx >= 0 ? this.storeIds.splice(idx, 1) : this.storeIds.push(store.ID);
    this.updateStoreSelection();
  }

  selectReportTotal(position: 'TOP' | 'BOTTOM') {
    this.selectedReportTotal = position;
    this.isReportDropdownOpen = false;
  }

  openGroupDetails(store: any, group: any) {
    this.selectedGroup = group;
    this.partsDetails = true;
  }

  closeGroupDetails() {
    this.selectedGroup = null;
    this.partsDetails = false;
  }

  applyFilters() {
    // 1️⃣ Filter by selected store IDs
    let filteredStores = this.originalStoreData; // keep a copy of full data
    if (this.storeIds.length > 0) {
      filteredStores = filteredStores.filter((store: { ID: number; }) => this.storeIds.includes(store.ID));
    }

    // 2️⃣ Filter by selected source
    if (this.selectedStoreSource) {
      filteredStores = filteredStores.map((store: { AgeGroupData: any[]; }) => {
        const filteredAgeGroups = store.AgeGroupData.filter((group: { Source: any; }) =>
          group.Source === this.selectedStoreSource // assuming AgeGroupData has Source field
        );
        return {
          ...store,
          AgeGroupData: filteredAgeGroups
        };
      });
    }

    // 3️⃣ Filter by report total (TOP / BOTTOM)
    if (this.selectedReportTotal) {
      filteredStores = filteredStores.map((store: { AgeGroupData: any; }) => {
        let sortedAgeGroups = [...store.AgeGroupData];
        sortedAgeGroups.sort((a, b) => b.Total_Qty - a.Total_Qty); // descending by total qty

        if (this.selectedReportTotal === 'TOP') {
          sortedAgeGroups = sortedAgeGroups.slice(0, 5); // top 5
        } else if (this.selectedReportTotal === 'BOTTOM') {
          sortedAgeGroups = sortedAgeGroups.slice(-5); // bottom 5
        }

        return {
          ...store,
          AgeGroupData: sortedAgeGroups
        };
      });
    }

    // Assign filtered data to main table
    this.storeData = filteredStores;

    // Close dropdowns after applying
    this.isSourceDropdownOpen = false;
    this.isReportDropdownOpen = false;
    this.activePopover = -1;
  }

  onApply(): void {
    this.getPartsAgingReport();
  }
  // this.api.postmethod

  getPartsAgingReport() {
    this.shared.spinner.show();

    const payload = {
      Store: '2', // Static store ID as per original code
      UserID: 0,
      Source: this.selectedStoreSources.length > 0 ? this.selectedStoreSources.join(',') : '',
      // DateType: this.DateType
    };


    this.shared.api
      .postmethod(this.shared.common.routeEndpoint + 'GetPartsAgingReport', payload).subscribe({
        next: (res: any) => {
          this.shared.spinner.hide();

          const responseData = res.response || res || [];
          const filteredData = responseData.filter((store: any) => store.StoreName !== 'REPORT TOTAL');

          // ✅ Parse each store’s AgeGroupData
          this.storeData = filteredData.map((store: any) => ({
            ...store,
            AgeGroupData: store.AgeGroupData ? JSON.parse(store.AgeGroupData) : []
          }));
          this.originalStoreData = [...this.storeData];

          if (this.selectedReportTotal == 'TOP') {
            let last = this.storeData.pop()
            this.storeData.unshift(last)
          }
        },
        error: () => {
          this.shared.spinner.hide();
        }
      });
  }
  // Helper to parse AgeGroupData string
  getAgeGroups(item: any) {
    try {
      return JSON.parse(item.AgeGroupData);
    } catch {
      return [];
    }
  }

  getPartsDetails(store: any, group: any) {
    this.partsDetails = true;

    const payload = {
      store_id: store.Store,
      daterange: group.Never_sold
    };


    this.shared.api
      .postmethod(this.shared.common.routeEndpoint + 'GetPartsAgingJSONDetails', payload).subscribe({
        next: (res: any) => {
          if (res?.status === 200 && res.response?.length > 0) {
            const data = res.response[0];
 const rawParts = JSON.parse(data.PartsJsonDetails || '[]') || [];

rawParts.forEach((p: { Manufacturer: any; }) => p.Manufacturer = data.Mfr || '-'); // assign Mfr first

this.groupedParts = rawParts.reduce((acc: any, item: any) => {
  const key = item.Manufacturer;
  acc[key] = acc[key] || [];
  acc[key].push(item);
  return acc;
}, {});
  console.log(this.groupedParts);
  
          } else {
            this.groupedParts = {};
          }
        },
        error: (err: any) => {
          console.error('Error fetching parts details:', err);
          this.partsDetailsList = [];
        }
      });
  }

  // Toggle dropdown
  toggleSourceDropdown() {
    this.isSourceDropdownOpen = !this.isSourceDropdownOpen;
  }
  selectedStoreSources: string[] = []; // store all selected sources

  toggleSource(store: any) {
    const value = store.AP_Source;
    const index = this.selectedStoreSources.indexOf(value);

    if (index >= 0) {
      this.selectedStoreSources.splice(index, 1);
    } else {
      this.selectedStoreSources.push(value);
    }
  }

  toggleAllSources() {
    if (this.selectedStoreSources.length === this.partsSourceList.length) {
      // Unselect all
      this.selectedStoreSources = [];
    } else {
      // Select all
      this.selectedStoreSources = this.partsSourceList.map(s => s.AP_Source);
    }
  }

  // Fetch parts source with static store IDs
  getPartsSource() {
    const payload = {
      StoreID: '2'

    };

    this.shared.api
      .postmethod(this.shared.common.routeEndpoint + 'GetPartsSource', payload).subscribe({
        next: (res: any) => {
          if (res?.status === 200 && res.response?.length > 0) {
            this.partsSourceList = res.response; // populate dropdown options
            if (this.partsSourceList.length > 0) {
              this.selectedStoreSources = this.partsSourceList.map(s => s.AP_Source); // select all by default
            }
          } else {
            this.partsSourceList = [];
          }
        },
        error: (err: any) => {
          console.error('Error fetching parts source:', err);
          this.partsSourceList = [];
        }
      });
  }


  calculateQtyPer(qty: number, total: number): number {
    if (!qty || !total || total === 0) return 0;
    return (qty * 100) / total;
  }

  calculatestock(store: any) {
    if (!store.Stock_Cost) return 0;
    return (store.Stock_Idle * 100) / store.Stock_Cost;
  }

  calculatenonstock(store: any) {
    if (!store.NonStock_Cost) return 0;
    return (store.NonStock_Idle * 100) / store.NonStock_Cost;
  }

  calculatetotal(store: any) {
    if (!store.Total_Cost) return 0;
    return (store.Total_Idle * 100) / store.Total_Cost;
  }


  exportToExcel() {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet("Parts Aging");

    const DATE_EXT = this.datepipe.transform(new Date(), "MMddyyyy");

    /* ---------------------- TITLE ---------------------- */
    worksheet.addRow(["Parts Aging Report"]);
    worksheet.mergeCells("A1:L1");
    worksheet.getRow(1).font = { size: 16, bold: true };
    worksheet.addRow([]);

    /* ---------------------- FILTERS ---------------------- */
    worksheet.addRow(["Filters"]);
    worksheet.getRow(3).font = { bold: true };

    worksheet.addRow(["Report Totals:", this.selectedReportTotal || "ALL"]);
    worksheet.addRow([
      "Selected Sources:",
      this.selectedStoreSources.length ? this.selectedStoreSources.join(", ") : "ALL",
    ]);

    worksheet.addRow(["Selected Stores:", 'WesternAuto']);
    worksheet.addRow(["Generated On:", this.datepipe.transform(new Date(), "MM/dd/yyyy h:mm a")]);

    worksheet.addRow([]);

    /* ---------------------- HEADERS (Client-A format) ---------------------- */
    const headers = [
      "AS OF",

      "STOCK QTY",
      "STOCK QTY %",
      "STOCK COST",
      "STOCK IDLE %",

      "NON-STOCK QTY",
      "NON-STOCK QTY %",
      "NON-STOCK COST",
      "NON-STOCK IDLE %",

      "TOTAL QTY",
      "TOTAL QTY %",
      "TOTAL COST",
      "TOTAL IDLE %",
    ];

    const headerRow = worksheet.addRow(headers);
    headerRow.font = { bold: true, color: { argb: "FFFFFF" } };

    headerRow.eachCell((cell) => {
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "1f4e78" },
      };
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
    });

    worksheet.columns = [
      { width: 25 },

      { width: 12 }, { width: 12 }, { width: 15 }, { width: 12 },
      { width: 12 }, { width: 12 }, { width: 15 }, { width: 12 },
      { width: 12 }, { width: 12 }, { width: 15 }, { width: 12 },
    ];

    /* ---------------------- STORE DATA ---------------------- */
    this.storeData.forEach((store) => {
      /* -------- STORE HEADER (bold, colored) -------- */
      const storeRow = worksheet.addRow([
        store.StoreName,

        store.Stock_Qty,
        this.calculateQtyPer(store.Stock_Qty, store.Total_Qty),
        store.Stock_Cost,
        this.calculatestock(store),

        store.NonStock_Qty,
        this.calculateQtyPer(store.NonStock_Qty, store.Total_Qty),
        store.NonStock_Cost,
        this.calculatenonstock(store),

        store.Total_Qty,
        this.calculateQtyPer(store.Total_Qty, store.Total_Qty),
        store.Total_Cost,
        this.calculatetotal(store),
      ]);

      storeRow.font = { bold: true, color: { argb: "FFFFFF" } };
      storeRow.eachCell((c) => {
        c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "2e74b5" } };
        c.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
      });

      /* -------- AGE GROUP ROWS -------- */
      store.AgeGroupData.forEach((grp: { Never_sold: any; Stock_Qty: number; Stock_Cost: any; Stock_Idle: number; NonStock_Qty: number; NonStock_Cost: any; NonStock_Idle: number; Total_Qty: number; Total_Cost: any; Total_Idle: number; }) => {
        const row = worksheet.addRow([
          grp.Never_sold,

          grp.Stock_Qty,
          this.calculateQtyPer(grp.Stock_Qty, store.Stock_Qty),
          grp.Stock_Cost,
          grp.Stock_Idle > 0 ? (grp.Stock_Idle * 100) / store.Stock_Cost : "-",

          grp.NonStock_Qty,
          this.calculateQtyPer(grp.NonStock_Qty, store.NonStock_Qty),
          grp.NonStock_Cost,
          grp.NonStock_Idle > 0 ? (grp.NonStock_Idle * 100) / store.NonStock_Cost : "-",

          grp.Total_Qty,
          this.calculateQtyPer(grp.Total_Qty, store.Total_Qty),
          grp.Total_Cost,
          grp.Total_Idle > 0 ? (grp.Total_Idle * 100) / store.Total_Cost : "-",
        ]);

        row.eachCell((cell) => {
          cell.border = {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" },
          };
        });
      });

      worksheet.addRow([]);
    });

    /* ---------------------- DOWNLOAD ---------------------- */
    workbook.xlsx.writeBuffer().then((buffer) => {
      FileSaver.saveAs(
        new Blob([buffer], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        }),
        "PartsAging_" + DATE_EXT + ".xlsx"
      );
    });
  }
}

