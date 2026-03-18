import { Component, HostListener } from '@angular/core';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { BsDatepickerConfig, BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { Stores } from '../../../../CommonFilters/stores/stores';
import { Subscription } from 'rxjs';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { common } from '../../../../common';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
import { DatePipe } from '@angular/common';
import { NgxSpinnerService } from 'ngx-spinner';
import { Api } from '../../../../Core/Providers/Api/api';
import { environment } from '../../../../../environments/environment';
import { NgbModal } from '@ng-bootstrap/ng-bootstrap';
import { Deal } from '../../../../Layout/cdpdataview/deal/deal-module';
import FileSaver from 'file-saver';
import { Workbook } from 'exceljs';
import { Title } from '@angular/platform-browser';
declare var bootstrap: any;

@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [SharedModule, BsDatepickerModule, Stores],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
})
export class Dashboard {

  DealType: any = ['Retail', 'Wholesale'];
  DealStatus: any = ['Finalized', 'Booked', 'Delivered'];
  ZeroDifference: any = 'Y';

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
  currentMonth: any = new Date();
  selectedDate: Date = new Date();
  bsConfig: Partial<BsDatepickerConfig> = {
    dateInputFormat: 'MMMM/YYYY',
    minMode: 'month',
    maxDate: new Date()
  };
  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent) {
    const clickedInside = (event.target as HTMLElement).closest('.dropdown-toggle, .reportstores-card , .timeframe');
    if (!clickedInside) {
      this.activePopover = -1;
    }
  }
  constructor(public shared: Sharedservice,
    private apiSrvc: Api,
    private ngbmodal: NgbModal,
    private comm: common,
    private toast: ToastService,
    private datepipe: DatePipe,
    private spinner: NgxSpinnerService,private title: Title) {
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
      this.title.setTitle(this.comm.titleName + '-Booked to Final Reconciliation');
      const data = {
        title: 'Booked to Final Reconciliation',
        stores: this.storeIds,

        count: 0,
      };
      this.shared.api.SetHeaderData({
        obj: data,
      });
      this.currentMonth = this.selectedDate;
      this.GetData(this.currentMonth);
    }
  }
  activePopover: number = -1;
  togglePopover(popoverIndex: number) {
    this.activePopover = this.activePopover === popoverIndex ? -1 : popoverIndex;
  }

  multipleorsingle(block: string, value: string) {

    if (block === 'ZD') {
      this.ZeroDifference = value;
    }

    if (block === 'DT') {
      const index = this.DealType.indexOf(value);
      if (index >= 0) {
        this.DealType.splice(index, 1);
      } else {
        this.DealType.push(value);
      }
    }

    if (block === 'DS') {
      const index = this.DealStatus.indexOf(value);
      if (index >= 0) {
        this.DealStatus.splice(index, 1);
      } else {
        this.DealStatus.push(value);
      }
    }
  }


  applyDateChange() {
    console.log('Deal Type', this.DealType.toString());
    console.log('Deal Status', this.DealStatus.toString());
    console.log('Zero Difference', this.ZeroDifference);
    if (!this.storeIds || this.storeIds.length === 0) {
      this.toast.show(
        'Please Select Atleast One Store',
        'warning',
        'Warning'
      );
      return;
    }
    if (!this.DealType || this.DealType.length === 0) {
      this.toast.show(
        'Please Select Atleast One Deal Type',
        'warning',
        'Warning'
      );
      return;
    }
    if (!this.DealStatus || this.DealStatus.length === 0) {
      this.toast.show(
        'Please Select Atleast One Deal Status',
        'warning',
        'Warning'
      );
      return;
    }
    this.currentMonth = this.selectedDate;
    this.GetData(this.currentMonth);
  }

  BookedReconciliattionData: any;
  BookedReconciliattionDataXlsx: any = [];
  NoData: boolean = false;
  totalDealFEG = 0;
  totalDealBEG = 0;
  totalDealTotal = 0;
  totalAcctFEG = 0;
  totalAcctBEG = 0;
  totalAcctTotal = 0;
  groupedRows: { [key: string]: any[] } = {};
  objectKeys = Object.keys;
  totalDiffFEG = 0;
  totalDiffBEG = 0;
  totalDiffTotal = 0;
  GetData(Date: any) {
    this.spinner.show();
    this.BookedReconciliattionData = [];
    let Obj = {
      AS_IDs: this.storeIds.toString(),
      Date: this.datepipe.transform(Date, 'yyyy-MM-dd'),
      IncludeZeros: this.ZeroDifference,
      DealType: this.DealType.toString(),
      DealStatus: this.DealStatus.toString()
    };
    // console.log(Obj);
    const curl = environment.apiUrl + this.comm.routeEndpoint + '/GetBookedFinalSalesReconcilition';
    this.apiSrvc.postmethod(this.comm.routeEndpoint + '/GetBookedFinalSalesReconcilition', Obj).subscribe(
      (res) => {
        const currentTitle = document.title;
        this.apiSrvc.logSaving(curl, {}, '', res.message, currentTitle);
        if (res.status == 200) {
          if (res.response != undefined) {
            if (res.response.length > 0) {
              this.spinner.hide();

              this.BookedReconciliattionData = res.response.map((v: any) => ({
                ...v, controlState: '+', notesView: '+'
              }));

              this.BookedReconciliattionData.some(function (x: any) {
                if (x.NOTES != undefined && x.NOTES != null) {
                  x.NOTES = JSON.parse(x.NOTES);
                  x.duplicateNotes = x.NOTES;
                  x.Notesstate = false;
                  // if (x.Notes.length > 2) {
                  //   x.duplicateNotes = x.duplicateNotes.slice(0, 3);
                  // } else {
                  //   x.duplicateNotes = x.Notes;
                  // }
                  // console.log(x.Notes, x.duplicateNotes);
                }
              });
              console.log('SR Data', this.BookedReconciliattionData);

              this.groupDealsByType(this.BookedReconciliattionData);

              this.BookedReconciliattionDataXlsx = [...res.response];
              this.totalDealFEG = this.getTotalSum(this.BookedReconciliattionData, 'Deal_FEG');
              this.totalDealBEG = this.getTotalSum(this.BookedReconciliattionData, 'Deal_BEG');
              this.totalDealTotal = this.getTotalSum(this.BookedReconciliattionData, 'Deal_Total');


              this.totalAcctFEG = this.getTotalSum(this.BookedReconciliattionData, 'Acct_FEG');
              this.totalAcctBEG = this.getTotalSum(this.BookedReconciliattionData, 'Acct_BEG');
              this.totalAcctTotal = this.getTotalSum(this.BookedReconciliattionData, 'Acct_Total');

              this.totalDiffFEG = this.getTotalSum(this.BookedReconciliattionData, 'Diff_FEG');
              this.totalDiffBEG = this.getTotalSum(this.BookedReconciliattionData, 'Diff_BEG');
              this.totalDiffTotal = this.getTotalSum(this.BookedReconciliattionData, 'Diff_Total');
            } else {
              this.toast.show(res.status, 'danger', 'Error');
              this.spinner.hide();
              this.NoData = true;
            }
          } else {
            this.toast.show(res.status, 'danger', 'Error');
            this.spinner.hide();
            this.NoData = true;
          }
        }
        else {
          this.NoData = true;
          this.spinner.hide();
        }
      },
      (error) => {
        console.log(error);
      }
    );
  }

  groupDealsByType(data: any[]) {
    this.groupedRows = data.reduce((acc, row) => {
      const key = row.Custom_DealType || 'Other Deals';
      if (!acc[key]) acc[key] = [];
      acc[key].push(row);
      return acc;
    }, {} as { [key: string]: any[] });
    console.log(this.groupedRows, 'ggggggrrrrrrrrrrrrr')
  }
  getTotalSum(data: any[], field: string): number {
    return data
      .map(item => item[field])
      .reduce((sum, value) => sum + value, 0);
  }

  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    }
    return false;
  }
  Bookedstate: any;
  ngAfterViewInit(): void {
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.shared.common.pageName == 'Booked to Final Reconciliation') {
        if (res.obj.storesData != undefined) {
          this.groupsArray = res.obj.storesData;
          this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
          this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_name : this.groupName = ''
          this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
          this.getStoresandGroupsValues()
        }
      }
    })
    this.shared.api.getExportToExcelAllReports().subscribe((res) => {
      this.Bookedstate = res.obj.state;
      if (res.obj.title == 'Booked to Final Reconciliation') {
        if (res.obj.state == true) {
          this.exportAsXLSX();
        }
      }
    });
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
      'type': 'S', 'others': 'N'
    };
  }

  callLoadingState = 'FL'
  toggleView(data: any) {
    if (data.notesView == '+') {
      data.notesView = '-'
    } else {
      data.notesView = '+'
    }
  }
  notesViewState: boolean = true
  notesView() {
    this.notesViewState = !this.notesViewState
  }
  Scrollpercent: any = 0;
  scrollCurrentposition: any = 0
  selecteddata: any = [];
  scrollpositionstoring: any = 0
  notesStageValue: any = ''
  notesStageText: any = ''
  notesStageValueGrid: any = '';
  notesstage: any = [];
  getDropDown(companyid: any) {
    const obj = {
      AssociatedReport: 'BookedDeal',
      CompanyID: companyid.Storeid,
    };
    this.apiSrvc
      .postmethod(this.comm.routeEndpoint + 'GetScheduleNoteStages', obj)
      .subscribe((res: any) => {
        if (res.status === 200 && res.response) {
          this.notesstage = res.response;
          // this.addNotes(companyid)
          console.log(this.notesstage)
          // this.notesStageValue =  "";
        } else {
          console.error('Error fetching dropdown data:', res);
        }
      });
  }
  onNotesStageChange(value: any) {
    console.log('Selected NS_ID:', value);
    // Additional logic can be added here if needed
  }
  addNotes(item: any) {
    this.getDropDown(item);
    this.scrollpositionstoring = this.scrollCurrentposition

    this.selecteddata = item;
    this.notesStageText = '';
    this.notesStageValue = ''
  }
  save() {
    let finalarray: any = []
    if (this.notesStageText == '' && this.notesStageValue == '') {
      if (this.notesStageText == '' && this.notesStageValue == '') {
        this.toast.show('Please enter notes or select the stage', 'warning', 'Warning');
      }
    }
    else {
      const obj = {
        "AS_ID": this.selecteddata.Storeid,
        "Account": "BookedToFinal",
        "Control": this.selecteddata.Deal_DealID,
        "Notes": this.notesStageText,
        "StageId": this.notesStageValue,
        "UserID":  JSON.parse(localStorage.getItem('userInfo')!)?.user_Info?.userid,
      }
      console.log(obj);
      this.apiSrvc.postmethod(this.comm.routeEndpoint + 'AddScheduleNotesAction', obj).subscribe((res: any) => {
        if (res.status == 200) {
          this.spinner.hide()
          this.toast.show('Notes Added Successfully', 'success', 'Success');
          this.notesViewState = true
          this.callLoadingState = 'ANS'
          const modalElement = document.getElementById('payplan');
          if (modalElement) {
            const modalInstance = bootstrap.Modal.getOrCreateInstance(modalElement);
            modalInstance.hide();
          }
          setTimeout(() => {
            document.body.classList.remove('modal-open');

            const backdrops = document.getElementsByClassName('modal-backdrop');
            while (backdrops.length > 0) {
              backdrops[0].parentNode?.removeChild(backdrops[0]);
            }
          }, 300);
          this.GetData(this.selectedDate)
        } else {
          this.toast.show('Something went wrong please try again', 'danger', 'Error')

        }
      })
    }
  }

  async openSalesModal(dealnumber: any, vin: any, storeid: any, stock: any, source: any, custno: any) {
    const module = await import('../../../../Layout/cdpdataview/deal/deal-module');
    const component = module.Deal;

    const modalRef = this.shared.ngbmodal.open(component, { size: 'xl', windowClass: 'connectedmodal' });
    modalRef.componentInstance.data = { dealno: dealnumber, vin: vin, storeid: storeid, stock: stock, source: source, custno: custno }; // Pass data to the modal component    
    modalRef.result.then((result) => {
      // this.topScroll()
      console.log(result); // Handle modal close result
    }, (reason) => {
      // this.topScroll()
      console.log(`Dismissed: ${reason}`); // Handle dismiss reason
    });
  }
  groupBy(array: any[], key: string): { [key: string]: any[] } {
    return array.reduce((result, currentValue) => {
      const groupKey = currentValue[key] || 'Unknown';
      if (!result[groupKey]) {
        result[groupKey] = [];
      }
      result[groupKey].push(currentValue);
      return result;
    }, {});
  }
  ExcelStoreNames: any = [];

  exportAsXLSX(): void {

    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet('Booked to Final Reconciliation');

    // =============================
    // FREEZE HEADER
    // =============================

    worksheet.views = [{
      state: 'frozen',
      ySplit: 14,
      topLeftCell: 'A15',
      showGridLines: false,
    }];

    // =============================
    // STORE FILTER VALUE
    // =============================

    const selectedStoreIds: string[] =
      this.storeIds?.length
        ? this.storeIds.map((id: any) => id.toString())
        : [];

    const storeValue = (this.stores || [])
      .filter((s: any) => selectedStoreIds.includes(s.ID.toString()))
      .map((s: any) => s.storename.trim())
      .join(', ');

    // =============================
    // TITLE
    // =============================

    worksheet.addRow('');
    const titleRow = worksheet.addRow(['Booked to Final Reconciliation']);
    titleRow.font = { name: 'Arial', size: 12, bold: true };
    worksheet.mergeCells('A2:Q2');

    worksheet.addRow('');
    worksheet.addRow([this.datepipe.transform(new Date(), 'MM/dd/yyyy h:mm:ss a')]);

    worksheet.addRow(['Report Filters :']).font = { bold: true };
    worksheet.addRow(['Summary Type :', 'Month Summary']);
    worksheet.addRow(['Month :', this.datepipe.transform(this.selectedDate, 'MMMM yyyy')]);
    worksheet.addRow(['Stores :', storeValue]);
    worksheet.addRow(['Deal Type :', this.DealType?.toString() || '-']);
    worksheet.addRow(['Deal Status :', this.DealStatus?.toString() || '-']);
    worksheet.addRow(['Zero Difference :', this.ZeroDifference == 'Y' ? 'Include' : 'Exclude']);

    worksheet.addRow('');

    // =============================
    // GROUP HEADER
    // =============================

    const groupRow = worksheet.addRow([
      '', '', '', '', '', '', '',
      'Transactional', '', '',
      'Accounting', '', '',
      'Difference', '', ''
    ]);

    worksheet.mergeCells(`A${groupRow.number}:G${groupRow.number}`);
    worksheet.mergeCells(`H${groupRow.number}:J${groupRow.number}`);
    worksheet.mergeCells(`K${groupRow.number}:M${groupRow.number}`);
    worksheet.mergeCells(`N${groupRow.number}:P${groupRow.number}`);

    for (let col = 1; col <= 16; col++) {
      const cell = groupRow.getCell(col);
      cell.alignment = { horizontal: 'center' };
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '2A91F0' }
      };
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
    }

    // =============================
    // COLUMN HEADER
    // =============================

    const headerRow = worksheet.addRow([
      'NOTES',
      'Store', 'Stock #', 'New/Used', 'Deal Status', 'Deal ID', 'Date',
      'Front Gross', 'Back Gross', 'Total Gross',
      'Front Gross', 'Back Gross', 'Total Gross',
      'Front Gross', 'Back Gross', 'Total Gross'
    ]);

    headerRow.eachCell(cell => {
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.alignment = { horizontal: 'center' };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '788494' }
      };
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
    });

    // =============================
    // DATA
    // =============================

    const grouped = this.groupBy(this.BookedReconciliattionDataXlsx, 'Custom_DealType');

    Object.keys(grouped).forEach(groupLabel => {

      // SECTION TITLE
      const typeRow = worksheet.addRow([groupLabel.toUpperCase()]);
      worksheet.mergeCells(`A${typeRow.number}:P${typeRow.number}`);

      for (let col = 1; col <= 16; col++) {
        const cell = typeRow.getCell(col);
        cell.font = { bold: true };
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'D9E1F2' }
        };
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      }

      // ROW DATA
      grouped[groupLabel].forEach((d: any) => {

        const formatMoney = (val: any) =>
          val == null || val === '' ? null : Number(val);

        const row = worksheet.addRow([
          d.NotesStatus === 'Y' ? '📝' : '',
          d.Store || '-',
          d.Stockno || '-',
          d.Dealtype || '-',
          d.Deal_Status || '-',
          d.Deal_DealID || '-',
          d.Deal_Date
            ? this.datepipe.transform(d.Deal_Date, 'MM.dd.yyyy')
            : '-',

          formatMoney(d.Deal_FEG),
          formatMoney(d.Deal_BEG),
          formatMoney(d.Deal_Total),

          formatMoney(d.Acct_FEG),
          formatMoney(d.Acct_BEG),
          formatMoney(d.Acct_Total),

          formatMoney(d.Diff_FEG),
          formatMoney(d.Diff_BEG),
          formatMoney(d.Diff_Total),
        ]);

        row.eachCell((cell: any, col: number) => {

          cell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
          };

          if (col >= 8) {
            cell.numFmt = '"$"#,##0';
            cell.alignment = { horizontal: 'right' };

            if (cell.value && Number(cell.value) < 0) {
              cell.font = { color: { argb: 'FF0000' } };
            }
          } else {
            cell.alignment = { horizontal: 'left' };
          }
        });

        // =============================
        // TOTAL ROW (same as UI logic)
        // =============================

        if (d.NotesStatus == null) {

          row.eachCell((cell: any) => {
            cell.font = { bold: true };

            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'EFEFEF' }
            };

            cell.border = {
              top: { style: 'medium' },
              left: { style: 'thin' },
              bottom: { style: 'medium' },
              right: { style: 'thin' }
            };
          });
        }

        // =============================
        // NOTES ROWS
        // =============================

        if (d.NOTES?.length) {

          d.NOTES.forEach((note: any) => {

            const noteRow = worksheet.addRow([
              '',
              'NOTES:',
              note.NOTES || '-'
            ]);

            worksheet.mergeCells(`C${noteRow.number}:P${noteRow.number}`);

            noteRow.font = { italic: true };

            noteRow.eachCell(cell => {
              cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
              };
            });

          });

        }

      });

    });

    // =============================
    // COLUMN WIDTH
    // =============================

    worksheet.getColumn(1).width = 20;
    worksheet.getColumn(2).width = 30;

    for (let i = 3; i <= 17; i++) {
      worksheet.getColumn(i).width = 15;
    }

    // =============================
    // AUTO FILTER
    // =============================

    worksheet.autoFilter = {
      from: { row: headerRow.number, column: 1 },
      to: { row: headerRow.number, column: 17 }
    };

    // =============================
    // SAVE FILE
    // =============================

    workbook.xlsx.writeBuffer().then((data: any) => {
      const blob = new Blob([data], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      });
      FileSaver.saveAs(blob, 'Booked to Final Reconciliation.xlsx');
    });

  }



}
