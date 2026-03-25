import { Component, OnInit, ViewChild, ElementRef, HostListener } from '@angular/core';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { Setdates } from '../../../../Core/Providers/SetDates/setdates';
import { BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { common } from '../../../../common';
import { Subscription } from 'rxjs';
import { environment } from '../../../../../environments/environment';
import { Stores } from '../../../../CommonFilters/stores/stores';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
import { CurrencyPipe } from '@angular/common';
import { NgbActiveModal } from '@ng-bootstrap/ng-bootstrap';
import { FilterPipe } from 'ngx-filter-pipe';

@Component({
  selector: 'app-dashboard',
  imports: [SharedModule, BsDatepickerModule, Stores],
    providers: [FilterPipe],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
})
export class Dashboard {

  TotalReport: any = 'T';
  AgeFrom: any = 21;
  AgeTo: any = 60;
  NoData: any;

  warrantytype: any = 'All';

  Pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;
  stores: any = []
  groupsArray: any = [];
  storename: any = ''
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  groupId: any = 0;
  storeIds: any = 0;
  MainGrid: any = 'Y'
  Role: any = 0
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
  constructor(
    public shared: Sharedservice, public setdates: Setdates, private comm: common, private cp: CurrencyPipe, private toast: ToastService, private ngmodelactive: NgbActiveModal
  ) {

    if (localStorage.getItem('flag') == 'V') {
      this.storeIds = [];
      console.log(JSON.parse(localStorage.getItem('userInfo')!), JSON.parse(localStorage.getItem('userInfo')!).user_Info, 'Widget Stores............');
      this.groupId = JSON.parse(localStorage.getItem('userInfo')!).groupid
      JSON.parse(localStorage.getItem('userInfo')!).store.indexOf(',') > 0 ?
        this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).wrstore.split(',') :
        this.storeIds.push(JSON.parse(localStorage.getItem('userInfo')!).wrstore)
      localStorage.setItem('flag', 'M')
    } else {
      if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
        this.groupId = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Preferences
        this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids.split(',')
        this.Role = JSON.parse(localStorage.getItem('userInfo')!).user_Info.roleid;

      }
    }
    if (this.shared.common.groupsandstores.length > 0) {
      this.groupsArray = this.shared.common.groupsandstores.filter((val: any) => val.sg_id != this.shared.common.reconID);
      this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
      this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_Name : this.groupName = ''
      this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
      this.getStoresandGroupsValues()
    }

    // this.storeIds = this.comm.redirectionFrom.flag == 'V' ? this.storeIds = this.comm.redirectionFrom.wrstore : this.commSrvc.StoresUserDetails('All', 0).toString();

    this.shared.setTitle(this.comm.titleName + "-Warranty Receivables");
    if (localStorage.getItem('childCalling') != undefined && localStorage.getItem('childCalling') != null) {
      let prevData = JSON.parse(localStorage.getItem('childDetailData')!)
      this.warrantytype = prevData[0].warrantytype;
      this.storeIds = prevData[0].globalstoreids;
      this.AgeFrom = prevData[0].age;
      this.AgeTo = prevData[0].to;
      this.groupId = prevData[0].group
    }
    this.setHeaderData()
    this.GetInventorySummaryReport();
  }
  newValuesToEdit: any = []
  isDesc: boolean = false;
  column: string = 'CategoryName';
  sort(property: any) {
    this.isDesc = !this.isDesc; //change the direction
    this.column = property;
    let direction = this.isDesc ? 1 : -1;
    this.InventorySummaryData.sort(function (a: any, b: any) {
      if (a[property] < b[property]) {
        return -1 * direction;
      } else if (a[property] > b[property]) {
        return 1 * direction;
      } else {
        return 0;
      }
    });
  }
  InventorySummaryData: any = [];
  inventory: any = []
  GetInventorySummaryReport() {
    localStorage.removeItem('childCalling')
    localStorage.removeItem('childDetailData')
    this.shared.spinner.show();
    const obj = {
      "RECEIVABLE_TYPE": this.warrantytype == 'All' ? '' : this.warrantytype,
      "AS_IDS": this.storeIds,
      "FW_AGE": this.AgeFrom,
      "EW_AGE": this.AgeTo,
      "UserID": this.comm.userId
    };
    // fredbeans/GetRecivablesSummary
    const curl = environment.apiUrl + this.comm.routeEndpoint + 'GetRecivablesSummary';
    this.shared.api.postmethod("fredbeans/GetRecivablesSummary", obj).subscribe(
      (totalres) => {
        const currentTitle = document.title;
        this.shared.api.logSaving(curl, {}, '', totalres.message, currentTitle);
        if (totalres.status == 200) {
          if (totalres.response != undefined) {
            // this.InventorySummaryData = totalres.response
            let data = totalres.response;
            data.some(function (x: any) {
              if (x.Comments != undefined && x.Comments != null) {
                x.Comments = JSON.parse(x.Comments);
              }
            });
            this.InventorySummaryData = data.reduce((r: any, { STORENAME, ...rest }: any) => {
              if (!r.some((o: any) => o.STORENAME == STORENAME)) {
                r.push({
                  STORENAME,
                  ...rest,
                  warranty: data.filter((v: any) => v.STORENAME == STORENAME),
                });
              }
              return r;
            }, []);
            //console.log(this.InventorySummaryData, '.................');
            this.inventory = this.InventorySummaryData
            let arr = [...this.InventorySummaryData]
            // this.newValuesToEdit.push(arr[0])
            this.newValuesToEdit = JSON.parse(JSON.stringify(arr))
          }
          this.shared.spinner.hide();
          if (this.InventorySummaryData.length > 0) {
            this.NoData = false;
          } else {
            this.NoData = true;
          }
        }
      })
  }
  closepopup() {
    this.ngmodelactive.dismiss()
    this.InventorySummaryData = this.inventory
  }
  getNewdetails(temp: any) {
    this.ngmodelactive = this.shared.ngbmodal.open(temp, {
      size: 'xl',
      backdrop: 'static',
    });
    this.newValuesToEdit.some((x: any) => {
      x.warranty.forEach((val: any, i: any) => {
        if (i != 0) {
          if (val.GUIDE == 0) {
            val.TargetValue = ''
          } else {
            val.TargetValue = this.addPercentageSymbol(val.GUIDE)
          }
        }
      });
    });
  }
  addPercentageSymbol(value: any) {
    let stringValue = String(value);
    if (stringValue.trim() !== '' && stringValue.trim() !== '0') {
      let numericValue = stringValue.replace(/[^0-9.]/g, '');
      let parsedValue = parseFloat(numericValue);
      if (isNaN(parsedValue)) {
        console.error('Value is not a valid number:', stringValue);
        return stringValue;
      }
      if (stringValue.trim().startsWith('$')) {
        return stringValue.trim();
      }
      let formattedValue = '$' + parsedValue;
      return formattedValue;
    } else {
      return stringValue;
    }
  }
  activePopover: number = -1;

  togglePopover(popoverIndex: number) {
    this.activePopover = this.activePopover === popoverIndex ? -1 : popoverIndex;
  }
  save() {
    //console.log(this.newValuesToEdit);
    let finalarray: any = []
    this.newValuesToEdit.forEach((val: any, i: any) => {
      val.warranty.forEach((ele: any, j: any) => {
        if (ele.TargetValue != '' && j != 0 && val.STORENAME != 'Report Total') {
          finalarray.push({
            "AS_ID": ele.STOREID,
            "RECEIVABLE_TYPE": ele.WARRANTY_TYPE,
            "VALUE": ele.TargetValue.substring(1),
          })
        }
      })
    })
    //console.log(finalarray);
    this.shared.spinner.show()
    const obj = {
      receivableguide: finalarray
    }
    //console.log(obj);
    this.shared.api.postmethod(this.comm.routeEndpoint + 'AddReceivablesGuide', obj).subscribe((res: any) => {
      if (res.status == 200) {
        this.shared.spinner.hide()
        this.toast.success('Guide values Inserted Successfully')
        this.GetInventorySummaryReport()
        // document.getElementById('close')?.click()
        this.closepopup()
      } else {
        this.shared.spinner.hide()
        this.toast.success('Something went wrong please try again')
      }
    })
  }
  CompleteComponentState: boolean = true;
  subdataindex: any = 0;
  warrantyData: any = []
  openDetails(Item: any) {
    console.log(Item);
    this.MainGrid = 'N'
    this.warrantyData = [{
      storeid: Item.STOREID == 0 ? this.storeIds : Item.STOREID,
      storename: Item.STORENAME,
      warrantytype: this.warrantytype,
      globalstoreids: this.storeIds,
      age: this.AgeFrom,
      group: this.groupId,
      to: this.AgeTo,
      displaywarrantytype: Item.WARRANTY_TYPE,
      // companyid:Item.COMPANYID
    }]
    this.getDropDown();
    this.getDropDownAction();
    this.GetWarrantyReceivablesDetails(this.warrantyData[0])

  }

  setHeaderData() {
    const data = {
      title: 'Warranty Receivables',
      stores: this.storeIds,
      toporbottom: this.TotalReport,
      AgeFrom: this.AgeFrom,
      AgeTo: this.AgeTo,
      groups: this.groupId,
      warrantytype: this.warrantytype,
    };
    this.shared.api.SetHeaderData({ obj: data });
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

  multipleorsingle(block: any, e: any) {
    if (block == 'WT') {
      if (e == 'All') {
        const index = this.warrantytype.findIndex((i: any) => i == e);
        if (index >= 0) {
          this.warrantytype = [];
        } else {
          this.warrantytype = [];
          this.warrantytype.push(e)
        }
      } else {
        this.warrantytype = [];
        this.warrantytype.push(e);
      }
    }
  }
  viewreport() {
    this.activePopover = -1
    if (this.storeIds.length == 0 || this.AgeFrom == 0 || this.AgeTo == 0) {
      if (this.AgeFrom == 0) {
        this.toast.show('Factory must be greaterthan 0', 'warning', 'Warning');
      }
      if (this.AgeTo == 0) {
        this.toast.show('Extended must be greaterthan 0', 'warning', 'Warning');
      }
      if (this.storeIds.length == 0) {
        this.toast.show('Please select atleast one store', 'warning', 'Warning');
      }
    }
    else {
      this.setHeaderData()
      this.GetInventorySummaryReport()
    }
  }
  keyPressNumbers(event: any, val?: any) {
    var charCode = event.which ? event.which : event.keyCode;
    if (charCode < 48 || charCode > 57) {
      event.preventDefault();
      return false;
    } else {
      return true;
    }
  }
  ngAfterViewInit(): void {
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.comm.pageName == 'Warranty Receivables') {
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
        if (res.obj.title == 'Warranty Receivables') {
          if (res.obj.state == true) {
             if (this.MainGrid == 'Y') {
              this.exportToExcel();
            } else {
              this.exportToExcelDetails()
            }
          }
        }
      }
    });
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res) => {
      if (this.email != undefined) {
        if (res.obj.title == 'Warranty Receivables') {
          if (res.obj.stateEmailPdf == true) {
            //  this.sendEmailData(res.obj.Email, res.obj.notes, res.obj.from);
          }
        }
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res) => {
      if (this.print != undefined) {
        if (res.obj.title == 'Warranty Receivables') {
          if (res.obj.statePrint == true) {
            //  this.GetPrintData();
          }
        }
      }
    });
    this.Pdf = this.shared.api.getExportToPDFAllReports().subscribe((res) => {
      if (this.Pdf != undefined) {
        if (res.obj.title == 'Warranty Receivables') {
          if (res.obj.statePDF == true) {
            // this.generatePDF();
          }
        }
      }
    });
  }

  ngOnDestroy() {
    this.unsubscribe();
  }
  unsubscribe() {

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
  index = '';
  ExcelStoreNames: any = [];
  exportToExcel() {
    let storeNames: any = [];
    let store = this.storeIds.split(',');
    storeNames = this.comm.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores.filter((item: any) =>
      store.some((cat: any) => cat === item.ID.toString())
    );
    if (store.length == this.comm.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores.length) {
      this.ExcelStoreNames = 'All Stores'
    } else {
      this.ExcelStoreNames = storeNames.map(function (a: any) {
        return a.storename;
      });
    }
    const ServiceData = this.InventorySummaryData.map((_arrayElement: any) =>
      Object.assign({}, _arrayElement)
    );
    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('Warranty Receivables');
    worksheet.views = [
      {
        showGridLines: false,
      },
    ];
    worksheet.addRow('');
    const titleRow = worksheet.addRow(['Warranty Receivables']);
    titleRow.eachCell((cell: any, number: any) => {
      cell.alignment = { indent: 1, vertical: 'top', horizontal: 'left' };
    });
    titleRow.font = { name: 'Arial', family: 4, size: 12, bold: true };
    titleRow.worksheet.mergeCells('A2', 'D2');
    worksheet.addRow('');
    // const PresentMonth = this.shared.datePipe.transform(this.FromDate, 'MMMM');
    // const PresentYear = this.shared.datePipe.transform(this.FromDate, 'yyyy');
    // const FromDate = this.shared.datePipe.transform(this.FromDate, 'dd');
    // const ToDate = this.shared.datePipe.transform(this.ToDate, 'dd');
    const DateToday = this.shared.datePipe.transform(
      new Date(),
      'MM/dd/yyyy h:mm:ss a'
    );
    const DATE_EXTENSION = this.shared.datePipe.transform(
      new Date(),
      'MMddyyyy'
    );
    worksheet.addRow([DateToday]).font = { name: 'Arial', family: 4, size: 9 };
    const ReportFilter = worksheet.addRow(['Report Controls :']);
    ReportFilter.font = { name: 'Arial', family: 4, size: 10, bold: true };

    const Stores = worksheet.addRow(['Stores :']);
    Stores.getCell(1).font = { name: 'Arial', family: 4, size: 9, bold: true };
    const Groups = worksheet.getCell('B8');
    Groups.value = 'Groups :';
    const groups = worksheet.getCell('D8');
    //console.log(this.groups);
    groups.value = this.comm.groupsandstores.filter((val: any) => val.sg_id == this.groupId.toString())[0].sg_name;
    groups.font = { name: 'Arial', family: 4, size: 9 };
    const Brands = worksheet.getCell('B9');
    Brands.value = 'Brands :';
    const brands = worksheet.getCell('D9');
    brands.value = '-';
    brands.font = { name: 'Arial', family: 4, size: 9 };
    const Stores1 = worksheet.getCell('B10');
    Stores1.value = 'Stores :';
    const stores1 = worksheet.getCell('D10');
    stores1.value = this.ExcelStoreNames == 0
      ? 'All Stores'
      : this.ExcelStoreNames == null
        ? '-'
        : this.ExcelStoreNames.toString().replaceAll(',', ', ');
    stores1.font = { name: 'Arial', family: 4, size: 9 };
    const Filters = worksheet.addRow(['Filters :']);
    Filters.getCell(1).font = { name: 'Arial', family: 4, size: 9, bold: true };
    const ROType = worksheet.getCell('B11');
    ROType.value = 'Warranty Type :';
    const rotype = worksheet.getCell('D11');
    rotype.value = this.warrantytype[0];
    rotype.font = { name: 'Arial', family: 4, size: 9 };
    const Department = worksheet.getCell('B12');
    Department.value = 'Older than :';
    const department = worksheet.getCell('D12');
    department.value = 'Factory : ' + this.AgeFrom + ' , ' + ' Extended :' + this.AgeTo;
    department.font = { name: 'Arial', family: 4, size: 9 };

    worksheet.addRow('');
    let Headings = [
      '',
      'Aged Balance',
      'Total Balance',
      'Guide',
      'Difference',
      '0-30 Days',
      '31-60 Days',
      '61-90 Days',
      '90+ Days',
    ];
    const headerRow = worksheet.addRow(Headings);
    headerRow.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: false,
      color: { argb: 'FFFFFF' },
    };
    headerRow.alignment = { indent: 1, vertical: 'top', horizontal: 'center' };
    headerRow.eachCell((cell: any, number: any) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '2a91f0' }, bgColor: { argb: 'white' }
      };
      cell.border = { right: { style: 'thin' } };
      cell.alignment = { vertical: 'top', horizontal: 'center' };
    });
    for (const d of ServiceData) {
      const Data1 = worksheet.addRow([
        d.STORENAME == '' ? '-' : d.STORENAME == null ? '-' : d.STORENAME,
        d['AGED BALANCE'] == '' ? '-' : d['AGED BALANCE'] == null ? '-' : d['AGED BALANCE'],
        d['TOTAL'] == '' ? '-' : d['TOTAL'] == null ? '-' : d['TOTAL'],
        d.GUIDE == ''
          ? '-'
          : d.GUIDE == null
            ? '-'
            : d.GUIDE,
        d['DIFF BALANCE'] == '' ? '-' : d['DIFF BALANCE'] == null ? '-' : d['DIFF BALANCE'],
        d['0-30 DAYS'] == '' ? '-' : d['0-30 DAYS'] == null ? '-' : d['0-30 DAYS'],
        d['31-60 DAYS'] == '' ? '-' : d['31-60 DAYS'] == null ? '-' : d['31-60 DAYS'],
        d['61-90 DAYS'] == ''
          ? '-'
          : d['61-90 DAYS'] == null
            ? '-'
            : d['61-90 DAYS'],
        d['90+ DAYS'] == ''
          ? '-'
          : d['90+ DAYS'] == null
            ? '-'
            : d['90+ DAYS'],
      ]);
      // Data1.outlineLevel = 1; // Grouping level 1
      Data1.font = { name: 'Arial', family: 4, size: 9 };
      Data1.getCell(1).alignment = {
        indent: 1,
        vertical: 'top',
        horizontal: 'left',
      };
      Data1.eachCell((cell: any, number: any) => {
        cell.border = { right: { style: 'thin' } };
        cell.numFmt = '$#,##0.00';
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
      if (d.warranty != undefined) {
        // for (const d1 of d.warranty) {
        d.warranty.forEach((d1: any, i: any) => {
          if (i != 0) {
            const Data2 = worksheet.addRow([
              d1.WARRANTY_TYPE == '' ? '-' : d1.WARRANTY_TYPE == null ? '-' : d1.WARRANTY_TYPE,
              d1['AGED BALANCE'] == ''
                ? '-'
                : d1['AGED BALANCE'] == null
                  ? '-'
                  : d1['AGED BALANCE'],
              d1['TOTAL'] == ''
                ? '-'
                : d1['TOTAL'] == null
                  ? '-'
                  : d1['TOTAL'],
              d1.GUIDE == ''
                ? '-'
                : d1.GUIDE == null
                  ? '-'
                  : d1.GUIDE,
              d1['DIFF BALANCE'] == '' ? '-' : d1['DIFF BALANCE'] == null ? '-' : d1['DIFF BALANCE'],
              d1['0-30 DAYS'] == '' ? '-' : d1['0-30 DAYS'] == null ? '-' : d1['0-30 DAYS'],
              d1['31-60 DAYS'] == '' ? '-' : d1['31-60 DAYS'] == null ? '-' : d1['31-60 DAYS'],
              d1['61-90 DAYS'] == ''
                ? '-'
                : d1['61-90 DAYS'] == null
                  ? '-'
                  : d1['61-90 DAYS'],
              d1['90+ DAYS'] == ''
                ? '-'
                : d1['90+ DAYS'] == null
                  ? '-'
                  : d1['90+ DAYS'],
            ]);
            Data2.outlineLevel = 1; // Grouping level 2
            Data2.font = { name: 'Arial', family: 4, size: 8 };
            Data2.getCell(1).alignment = {
              indent: 2,
              vertical: 'top',
              horizontal: 'left',
            };
            Data2.eachCell((cell: any, number: any) => {
              cell.border = { right: { style: 'thin' } };
              cell.numFmt = '$#,##0.00';
              if (number != 1) {
                cell.alignment = {
                  vertical: 'top',
                  horizontal: 'right',
                  indent: 1,
                };
              }
            });
            if (Data2.number % 2) {
              Data2.eachCell((cell, number) => {
                cell.fill = {
                  type: 'pattern',
                  pattern: 'solid',
                  fgColor: { argb: 'e5e5e5' },
                  bgColor: { argb: 'FF0000FF' },
                };
              });
            }
          }
        })
        // }
      }
    }
    worksheet.eachRow((row, rowIndex) => {
      row.eachCell((cell, colIndex) => {
        if (rowIndex > 1 && rowIndex < 13) { // Skip the header row
          // Apply conditional alignment based on your conditions
          if (colIndex === 1) {
            // Apply right alignment to the second column
            cell.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
          }
        }
      });
    });
    worksheet.getColumn(1).width = 30;
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
    worksheet.getColumn(14).width = 15;
    worksheet.getColumn(15).width = 15;
    worksheet.getColumn(16).width = 15;
    worksheet.getColumn(17).width = 15;
    worksheet.getColumn(18).width = 15;
    worksheet.getColumn(19).width = 15;
    worksheet.getColumn(20).width = 15;
    worksheet.getColumn(21).width = 15;
    worksheet.addRow([]);

    this.shared.exportToExcel(workbook, 'Warranty Receivables_' + DATE_EXTENSION);


  }
  NoDataDetails: any = '';
  
  WarrantyReceivablesDetailsData: any = []
  SGDSearchName: any = '';

  callLoadingState = 'FL';
  notesStageValue: any = ''
  notesStageValueGrid: any = ''
  notesStageText: any = '';
  notesstage: any = []

  hideVisibility: boolean = false;
  collectHidevalues(e: any, val: any, confirmtemplate: any, ref: any) {
    if (ref == 'multi') {
      if (this.hideRecords.length == 0) {
        this.toast.show('Please select atleast one record to hide','warning','Warning');
        var element = <HTMLInputElement>document.getElementById('symbol');
        element.checked = false;
      } else {
        if (e.target.checked) {
          this.hideVisibility = true;
          this.ngmodelactive = this.shared.ngbmodal.open(confirmtemplate, {
            size: 'sm',
            backdrop: 'static',
          });
        }
      }
    } else {
      if (e.target.checked) {
        this.hideVisibility = true;
        this.hideRecords.push(val);
        // console.log(this.hideRecords);
      } else {
        const index = this.hideRecords.findIndex(
          (list: { CONTROL: any }) => list.CONTROL == val.CONTROL
        );
        this.hideRecords.splice(index, 1);
        // console.log(this.hideRecords);
      }
    }
  }
backtoWR(){
  this.MainGrid = 'Y'
}
  hideRecords: any = [];
  FinalArray: any = [];

  hideAdd() {
    console.log(this.hideRecords);
    // (<HTMLInputElement>document.getElementById('closeadd')).click();

    this.FinalArray = [];
    for (var i = 0; i < this.hideRecords.length; i++) {
      this.FinalArray.push({
        Receivable_Type: 'WarrantyAR',
        CompanyID: this.hideRecords[i].COMPANYID,
        AS_ID: this.hideRecords[i].storeid,

        Account: this.hideRecords[i].ACCOUNT,


        Control: this.hideRecords[i].CONTROL,
        Stock: '',
        Deal: '',
        Control_Status: 'Y',
        UserID: JSON.parse(localStorage.getItem('UserDetails')!).userid,
      });
    }
    if (this.FinalArray.length == 0) {
      this.toast.show('Please select atleast one record to hide','warning','Warning');
    }
    else {
      const obj = {
        receivableexcludecontrol: this.FinalArray,
      };
      console.log(obj);
      this.shared.api.postmethod('ReceivableExcludeControls', obj).subscribe(
        (res) => {
          if (res.status == 200) {
            this.toast.success('This Account Hidden Successfully ');
            (<HTMLInputElement>document.getElementById('closeadd')).click();

            // this.onclose();
            // this.ngbmodalActive.close();

            this.GetWarrantyReceivablesDetails(this.warrantyData[0]);
            this.hideRecords = [];
            this.hideVisibility = false;
          } else {
            this.toast.show(res.status, 'danger','Error');
            this.shared.spinner.hide();
            this.NoDataDetails = '';
          }
        },
        (error) => {
          this.toast.show('502 Bad Gate Way Error', 'danger','Error');
          this.shared.spinner.hide();
          this.NoDataDetails = '';
        }
      );
    }
  }

  onclose() {
    var element = <HTMLInputElement>document.getElementById('symbol');
    element.checked = false;
    this.shared.ngbmodal.dismissAll();
  }
  completedata: any = []
  GetWarrantyReceivablesDetails(data: any) {
    this.SGDSearchName = ''
    this.WarrantyReceivablesDetailsData = []
    this.shared.spinner.show();
    const obj = {
      "AS_ID": data.storeid,
      "RECEIVABLE_TYPE": data.displaywarrantytype,
      "STAGE": this.notesStageValueGrid
    };
    this.shared.api.postmethod("fredbeans/GetRecivablesDetails", obj).subscribe(
      (totalres) => {
        if (totalres.status == 200) {
          if (totalres.response != undefined) {
            this.WarrantyReceivablesDetailsData = totalres.response.map((v: any) => ({
              ...v, controlState: '+', notesView: '+'
            }));
            this.WarrantyReceivablesDetailsData.some(function (x: any) {
              if (x.NOTES != undefined && x.NOTES != '' && x.NOTES != null) {
                x.NOTES = JSON.parse(x.NOTES);
              }
              if (x.Comments != undefined && x.Comments != null) {
                x.Comments = JSON.parse(x.Comments);
              }
              if (x.TOTALS != undefined && x.TOTALS != null) {
                x.TOTALS = JSON.parse(x.TOTALS);
              }

            });
            this.callLoadingState == 'ANS' ? this.sortDetails(this.column, this.callLoadingState) : '';
            let position = this.scrollpositionstoring + 10
            setTimeout(() => {
              this.scrollcent.nativeElement.scrollTop = position
            }, 500);
          }
          this.completedata = this.WarrantyReceivablesDetailsData

          console.log('Warranty Receivables Data Details', this.WarrantyReceivablesDetailsData);
          this.shared.spinner.hide();
          if (this.WarrantyReceivablesDetailsData.length > 0) {
            this.NoDataDetails = '';
          } else {
            this.NoDataDetails = 'No Data Found';
          }
        } else {
          this.shared.spinner.hide();
          this.NoDataDetails = 'No Data Found';
        }
      })
  }
  filterData() {
    if (this.SGDSearchName.trim() !== '') {

      const search = this.SGDSearchName.toLowerCase();
      this.WarrantyReceivablesDetailsData = this.completedata
        .map((item: any) => {
          const matchesstorename = item.STORENAME?.toLowerCase().includes(search);
          const matchesaccount = item.ACCOUNT?.toLowerCase().includes(search);
          const matchesstage = item.STAGE?.toLowerCase().includes(search);
          const matchescontrol = typeof item.CONTROL === 'string' && item.CONTROL.toLowerCase().includes(search);
          const matchesbalance = typeof item.BALANCE === 'string' && item.BALANCE.toLowerCase().includes(search);
          const matchesage = typeof item.AGE === 'string' && item.AGE.toLowerCase().includes(search);

          const matchesOtherFields =
            matchesstorename || matchesaccount || matchesstage ||
            matchescontrol || matchesbalance || matchesage;

          if (matchesOtherFields) {
            return item; // Include as-is (keep all Notes)
          }

          // No primary field matched; check Notes
          if (Array.isArray(item.NOTES)) {
            const filteredNotes = item.NOTES.filter(
              (note: any) => note?.STAGE_NOTES?.toLowerCase().includes(search)
            );

            if (filteredNotes.length > 0) {
              return {
                ...item,
                NOTES: filteredNotes // only matched notes
              };
            }
          }

          // Neither fields nor Notes matched — exclude
          return null;
        })
        .filter((item: any) => item !== null);
    } else {
      this.WarrantyReceivablesDetailsData = [...this.completedata];
    }
    this.WarrantyReceivablesDetailsData && this.WarrantyReceivablesDetailsData.length > 0 ? this.NoData = '' : this.NoData = 'No Data Found';
    let position = this.scrollpositionstoring + 10
    setTimeout(() => {
      this.scrollcent.nativeElement.scrollTop = position
      // //console.log(position);

    }, 500);

  }
  getTotal(columnname: any, block: any) {
    let total = 0
    if (block == 'T') {
      this.WarrantyReceivablesDetailsData.some(function (x: any) {
        total += x[columnname]
      })
    }
    if (block == 'A') {
      let ageddata = this.WarrantyReceivablesDetailsData.filter((val: any) => val.AGE > 20)
      ageddata.some(function (x: any) {
        total += x[columnname]
      })
    }
    return total
  }

  controlData(item: any, sign: any) {
    if (sign == '+') {
      const obj = {
        "AS_ID": this.warrantyData[0].storeid,
        "ACCOUNT": item.ACCOUNT,
        "CONTROL": item.CONTROL
      };
      this.shared.api.postmethod("fredbeans/GetRecivablesControlEntries", obj).subscribe(
        (totalres) => {
          if (totalres.status == 200) {
            item.controlList = totalres.response;
            this.spinnerLoader = false;
            item.controlState = '-'
          }
        })
    } else {
      item.controlState = '+'
    }
  }
  greaterthanvalue: any = false;
  spinnerLoader: boolean = false;
  agegreaterthan(e: any) {
    console.log(e, 'e', this.greaterthanvalue);
    this.spinnerLoader = true;
    this.filterData()
    if (this.greaterthanvalue == true) {
      this.WarrantyReceivablesDetailsData = this.WarrantyReceivablesDetailsData.filter((val: any) => val.AGE > 89)
      this.spinnerLoader = false;
      // this.getTotal('BALANCE', 'T')
    } else {
      this.WarrantyReceivablesDetailsData = this.WarrantyReceivablesDetailsData;
      this.spinnerLoader = false;
      // this.getTotal('BALANCE', 'T')
    }
    this.WarrantyReceivablesDetailsData && this.WarrantyReceivablesDetailsData.length > 0 ? this.NoData = '' : this.NoData = 'No Data Found';

  }
  getDropDown() {

    const obj = {
      "AssociatedReport": 'WarrantyAR',
      "CompanyID": this.warrantyData[0].storeid,
      "Receivable_Type": this.warrantyData[0].displaywarrantytype
    }
    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetScheduleNoteStages', obj).subscribe((res: any) => {
      if (res.status == 200) {
        this.notesstage = res.response;
        this.notesStageValue = ''
      }
    })
  }
  getDropDownAction() {
    const obj = {
      "AssociatedReport": 'WarrantyAR',
      "CompanyID": this.warrantyData[0].storeid,
      "Receivable_Type": ''
    }
    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetScheduleNoteStages', obj).subscribe((res: any) => {
      if (res.status == 200) {
        this.notesstageAction = res.response;
        this.notesStageValue = ''
      }
    })
  }
  selecteddata: any = [];
  notesstageAction: any = []
  addNotes(item: any) {
    this.scrollpositionstoring = this.scrollCurrentposition
    this.selecteddata = item;
    this.notesStageText = '';
    this.notesStageValue = '';
  }
  sortDetails(property: any, state?: any) {
    if (state == undefined) {
      this.isDesc = !this.isDesc;
    }
    this.callLoadingState = 'FL'
    this.column = property;
    let direction = this.isDesc ? 1 : -1;
    this.WarrantyReceivablesDetailsData.sort(function (a: any, b: any) {
      if (a[property] < b[property]) {
        return -1 * direction;
      } else if (a[property] > b[property]) {
        return 1 * direction;
      } else {
        return 0;
      }
    });
  }
  toggleView(data: any) {
    if (data.notesView == '+') {
      data.notesView = '-'
    } else {
      data.notesView = '+'
    }
  }
    public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    }
    return false;
  }
  saveNotes() {
    if (this.notesStageText == '' || this.notesStageValue == '') {
      if (this.notesStageText == '') {
        this.toast.show('Please enter notes','warning','Warning');
      }
      if (this.notesStageValue == '') {
        this.toast.show('Please select Stage','warning','Warning');
      }
    }
    else {
      const obj = {

        "AS_ID": this.selecteddata.storeid,
        "Account": this.selecteddata.ACCOUNT,
        "Control": this.selecteddata.CONTROL,
        "Notes": this.notesStageText,
        "StageId": this.notesStageValue,
        "UserID"
          :
          JSON.parse(localStorage.getItem('UserDetails')!).userid
      }
      //console.log(obj);
      this.shared.api.postmethod(this.comm.routeEndpoint + 'AddScheduleNotesAction', obj).subscribe((res: any) => {
        if (res.status == 200) {
          this.shared.spinner.hide()
          this.toast.success('Notes Added Successfully');
          document.getElementById('close')?.click();
          this.callLoadingState = 'ANS'
          this.GetWarrantyReceivablesDetails(this.warrantyData[0])

        } else {
          this.toast.success('Something went wrong please try again')
        }
      })
    }
  }
  notesViewState: boolean = true
  notesView() {
    this.notesViewState = !this.notesViewState
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

  // ExcelStoreNames: any = [];
  exportToExcelDetails() {
    let storeNames: any = [];
    const ServiceData = this.WarrantyReceivablesDetailsData.map((_arrayElement: any) =>
      Object.assign({}, _arrayElement)
    );
    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('Warranty Receivables Details');
    worksheet.views = [
      {
        showGridLines: false,
      },
    ];
    worksheet.addRow('');
    const titleRow = worksheet.addRow(['Warranty Receivables Details']);
    titleRow.eachCell((cell: any, number: any) => {
      cell.alignment = { indent: 1, vertical: 'top', horizontal: 'left' };
    });
    titleRow.font = { name: 'Arial', family: 4, size: 12, bold: true };
    titleRow.worksheet.mergeCells('A2', 'D2');
    worksheet.addRow('');
    const DateToday = this.shared.datePipe.transform(
      new Date(),
      'MM/dd/yyyy h:mm:ss a'
    );
    const DATE_EXTENSION = this.shared.datePipe.transform(
      new Date(),
      'MMddyyyy'
    );
    worksheet.addRow([DateToday]).font = { name: 'Arial', family: 4, size: 9 };
    const ReportFilter = worksheet.addRow(['Report Controls :']);
    ReportFilter.font = { name: 'Arial', family: 4, size: 10, bold: true };
    const Stores = worksheet.addRow(['Stores :']);
    Stores.getCell(1).font = { name: 'Arial', family: 4, size: 9, bold: true };
    const Stores1 = worksheet.getCell('B10');
    Stores1.value = 'Stores :';
    const stores1 = worksheet.getCell('D10');
    stores1.value = this.warrantyData[0].storename;
    stores1.font = { name: 'Arial', family: 4, size: 9 };
    const Filters = worksheet.addRow(['Filters :']);
    Filters.getCell(1).font = { name: 'Arial', family: 4, size: 9, bold: true };
    const ROType = worksheet.getCell('B11');
    ROType.value = 'Warranty Type :';
    const rotype = worksheet.getCell('D11');
    rotype.value = this.warrantyData[0].displaywarrantytype;
    rotype.font = { name: 'Arial', family: 4, size: 9 };
    worksheet.addRow('');
    let Headings = [
      'Age',
      'Date',
      'Account #',
      'Stage',
      'Control',
      'Balance'
    ];
    const headerRow = worksheet.addRow(Headings);
    headerRow.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: false,
      color: { argb: 'FFFFFF' },
    };
    headerRow.alignment = { indent: 1, vertical: 'top', horizontal: 'center' };
    headerRow.eachCell((cell: any, number: any) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '2a91f0' }, bgColor: { argb: 'white' }
      };
      cell.border = { right: { style: 'thin' } };
      cell.alignment = { vertical: 'top', horizontal: 'center' };
    });
    worksheet.addRow('');
    let HeadingsBalance = [
      '',
      'Total Balance',
      this.getTotal('BALANCE', 'T'),
      '',
      'Aged Balance',
      this.getTotal('BALANCE', 'A'),
      ,
    ];
    const headerRowBal = worksheet.addRow(HeadingsBalance);
    headerRow.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: false,
      color: { argb: 'FFFFFF' },
    };
    headerRowBal.alignment = { indent: 1, vertical: 'top', horizontal: 'center' };
    headerRow.eachCell((cell: any, number: any) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '2a91f0' }, bgColor: { argb: 'white' }
      };
      cell.border = { right: { style: 'thin' } };
      cell.alignment = { vertical: 'top', horizontal: 'center' };
    });
    let count = 15
    for (const d of ServiceData) {
      count++
      const Data1 = worksheet.addRow([
        d.AGE == '' ? '-' : d.AGE == null ? '-' : d.AGE,
        d.DATE == '' ? '-' : d.DATE == null ? '-' : this.shared.datePipe.transform(d.DATE, 'MM.dd.YYYY'),
        d.ACCOUNT == '' ? '-' : d.ACCOUNT == null ? '-' : d.ACCOUNT,
        d.STAGE == '' ? '-' : d.STAGE == null ? '-' : d.STAGE,
        d.CONTROL == '' ? '-' : d.CONTROL == null ? '-' : d.CONTROL,
        d.BALANCE == '' ? '-' : d.BALANCE == null ? '-' : d.BALANCE,
      ]);
      // Data1.outlineLevel = 1; // Grouping level 1
      Data1.font = { name: 'Arial', family: 4, size: 9 };
      Data1.getCell(1).alignment = {
        indent: 1,
        vertical: 'top',
        horizontal: 'left',
      };
      Data1.eachCell((cell: any, number: any) => {
        cell.border = { right: { style: 'thin' } };
        if (number == 6) {
          cell.numFmt = '$#,##0.00';
        }
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
      if (d.NOTES != undefined && d.NOTES != null && this.notesViewState) {
        worksheet.mergeCells(count, 1, count, 6);
        const Data2NOtes = worksheet.getCell(count, 1);
        Data2NOtes.value = 'Notes'
        Data2NOtes.alignment = { indent: 2, vertical: 'middle', horizontal: 'left', };
        Data2NOtes.font = { name: 'Arial', family: 4, size: 9 };
        Data2NOtes.border = { right: { style: 'thin' }, left: { style: 'thin' }, top: { style: 'thin' }, bottom: { style: 'thin' } };
        count++
        for (const d1 of d.NOTES) {
          // worksheet.addRow([])
          worksheet.mergeCells(count, 1, count, 6);
          const Data2 = worksheet.getCell(count, 1);
          Data2.value = d1.STAGE_NOTES
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
    worksheet.eachRow((row, rowIndex) => {
      row.eachCell((cell, colIndex) => {
        if (rowIndex > 1 && rowIndex < 13) { // Skip the header row
          // Apply conditional alignment based on your conditions
          if (colIndex === 1) {
            // Apply right alignment to the second column
            cell.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
          }
        }
      });
    });
    worksheet.getColumn(1).width = 10;
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
    worksheet.getColumn(14).width = 15;
    worksheet.getColumn(15).width = 15;
    worksheet.getColumn(16).width = 15;
    worksheet.getColumn(17).width = 15;
    worksheet.getColumn(18).width = 15;
    worksheet.getColumn(19).width = 15;
    worksheet.getColumn(20).width = 15;
    worksheet.getColumn(21).width = 15;
    worksheet.addRow([]);
 
      this.shared.exportToExcel(workbook, 'Warranty Receivables Details_' + DATE_EXTENSION );
 
  }




}
