import { Component, OnInit, ViewChild, ElementRef, HostListener } from '@angular/core';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { Setdates } from '../../../../Core/Providers/SetDates/setdates';
import { Subscription } from 'rxjs';
import { Router } from '@angular/router'
import { NgbActiveModal } from '@ng-bootstrap/ng-bootstrap';
import { Stores } from '../../../../CommonFilters/stores/stores';
@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [SharedModule,Stores],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss'
})
export class Dashboard {
  storeIds: any = '0';
  storeName: any = 'CADILLAC BUICK GMC';
  NoData: any = '';
  groups: any = 1;
  MainGrid: any = 'Y'
  Block: any = '';
  var3: any = '';
  var2: any = '';
  ageCategory: any = '';
  DetailsNoData: any = ''
  inventoryBookDetails: any = []
  groupPreference: any = 30;
  Pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;
  userDetails: any;
  groupsArray: any = [];
  storename: any = ''
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  groupId: any = 8;
  stores: any = []
  storesFilterData: any = {
    'groupsArray': this.groupsArray, 'groupId': this.groupId, 'storesArray': this.stores, 'storeids': '1', 'type': 'M', 'others': 'N',
    'groupName': this.groupName, 'storename': this.storename, storecount: null, 'storedisplayname': this.storedisplayname
  };
  constructor(public shared: Sharedservice, public setdates: Setdates, private router: Router, public ngbModalActive: NgbActiveModal
  ) {
    if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
      this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.ustores.split(',')
    }
    if (this.shared.common.groupsandstores.length > 0) {
      this.groupsArray = this.shared.common.groupsandstores.filter((val: any) => val.sg_id != this.shared.common.reconID);
      this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
      this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_Name : this.groupName = ''
      this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
      this.getStoresandGroupsValues()
    }
    this.shared.setTitle(this.shared.common.titleName + '-Inventory Book');
    this.setHeaderData()
    this.GetInventorySummaryReport();
  }
  isDesc: boolean = false;
  column: string = 'CategoryName';
  newValuesToEdit: any = [];
  ngOnInit(): void { }
  setHeaderData() {
    const headerdata = {
      title: 'Inventory Book',
      stores: this.storeIds,
      groups: this.groups,
      stocktype: this.stockType,
      status: this.status,
      wholesale: this.wholesale,
      aged: this.aged,
    };
    this.shared.api.SetHeaderData({
      obj: headerdata,
    });
  }
  InventorySummaryData: any = [];
  key: any = 'data2'
  order: any = 'asc'
  sort(key: any, order: any) {
    if (this.key == key) {
      if (order == 'asc') {
        this.order = 'desc'
      } else {
        this.order = 'asc'
      }
    } else {
      this.order = 'asc'
    }
    this.key = key
    this.GetInventorySummaryReport()
  }
  Detailsort(property: any, state?: any) {
    if (state == undefined) {
      this.isDesc = !this.isDesc;
    }
    // this.callLoadingState = 'FL' //change the direction
    this.column = property;
    let direction = this.isDesc ? 1 : -1;
    // // console.log(property)
    this.inventoryBookDetails.sort(function (a: any, b: any) {
      if (a[property] < b[property]) {
        return -1 * direction;
      } else if (a[property] > b[property]) {
        return 1 * direction;
      } else {
        return 0;
      }
    });
  }
  GetInventorySummaryReport() {
    localStorage.removeItem('childCalling')
    localStorage.removeItem('childDetailData')
    this.shared.spinner.show();
    const obj = {
      as_id: this.storeIds,
      Header: this.key,
      Order: this.order,
      StockType: this.stockType.toString(),
      Status: this.status.toString(),
      WholeSale: this.wholesale.toString(),
      Type: this.aged.toString()
    };
    const curl = this.shared.getEnviUrl() + this.shared.common.routeEndpoint + 'GetInventoryBookV1';
    this.shared.api
      .postmethod(this.shared.common.routeEndpoint + 'GetInventoryBookV1', obj)
      .subscribe((totalres: { message: any; status: number; response: any; }) => {
        const currentTitle = document.title;
        this.shared.api.logSaving(curl, {}, '', totalres.message, currentTitle);
        if (totalres.status == 200) {
          this.InventorySummaryData = totalres.response;
          if (this.InventorySummaryData != undefined) {
            if (this.InventorySummaryData.length > 0) {
              this.InventorySummaryData.some(function (x: any) {
                if (x.data2Json != undefined) {
                  x.data2 = JSON.parse(x.data2Json);
                  x.data2 = x.data2.map((v: any) => ({
                    ...v,
                    SubData: [],
                    data2sign: '-',
                  }));
                }
                if (x.data3Json != undefined) {
                  x.data3 = JSON.parse(x.data3Json);
                  x.data2.forEach((val: any) => {
                    x.data3.forEach((ele: any) => {
                      if (val.data2 == ele.data2) {
                        val.SubData.push(ele);
                      }
                    });
                  });
                }
                x.Dealer = '+';
              });
            }
            console.log('Inventory Book Data', this.InventorySummaryData);
            let arr = [...this.InventorySummaryData];
            this.newValuesToEdit = JSON.parse(JSON.stringify(arr));
          }
          console.log('Inventory Resp : ', this.InventorySummaryData);
          this.shared.spinner.hide();
          if (this.InventorySummaryData.length > 0) {
            this.NoData = '';
          } else {
            this.NoData = 'No Data Found!!';
          }
        } else {
          this.shared.spinner.hide()
          this.NoData = 'No Data Found!!'
        }
      });
  }
  getTotal(columnname: any, block?: any) {
    let total = 0
    let i = 0
    this.InventorySummaryData[0].data3.some(function (x: any) {
      i++
      total += x[columnname]
    })
    if (block == 'AVG') {
      console.log(i);
      return total / i
    } else {
      return total
    }
  }
  getNewdetails(temp: any) {
    this.ngbModalActive = this.shared.ngbmodal.open(temp, {
      size: 'xl',
      backdrop: 'static',
    });
    this.newValuesToEdit.some(function (x: any) {
      x.data2.forEach((val: any) => {
        val.gridData = [];
        x.data3.forEach((ele: any) => {
          if (val.data2 == ele.data2) {
            if (ele.Target == 0 && ele.TargetByModel == '') {
              ele.TargetValue = '';
            } else {
              ele.TargetValue = ele.Target;
            }
            if (ele.TargetByModel_60 == 0 || ele.TargetByModel_60 == '') {
              ele.TargetValue60 = '';
            } else {
              ele.TargetValue60 = ele.TargetByModel_60;
            }
            if (ele.TargetByModel_90 == 0 || ele.TargetByModel_90 == '') {
              ele.TargetValue90 = '';
            } else {
              ele.TargetValue90 = ele.TargetByModel_90;
            }
            val.gridData.push(ele);
          }
        });
      });
    });
  }
  save() {
    let finalarray: any = [];
    this.newValuesToEdit[0].data2.forEach((val: any, i: any) => {
      var stkTyp = (this.stockType[0] == 'New' ? 'N' : (this.stockType[0] == 'Used' ? 'U' : 'F'))
      val.gridData.forEach((ele: any, j: any) => {
        console.log('R1 : ', ele.TargetValue);
        console.log('R2 : ', ele.TargetValue60);
        console.log('R3 : ', ele.TargetValue90);
        if (ele.Target !== ele.TargetValue) {
          finalarray.push({
            StoreID: ele.StoreID,
            ModelID: ele.ModelCode,
            StockType: stkTyp,
            Target: parseInt(ele.TargetValue),
            Target_60: parseInt(ele.TargetValue60),
            Target_90: parseInt(ele.TargetValue90),
            Status: 'Y',
          });
        } else if (ele.TargetByModel_60 !== ele.TargetValue60) {
          finalarray.push({
            StoreID: ele.StoreID,
            ModelID: ele.ModelCode,
            StockType: stkTyp,
            Target: parseInt(ele.TargetValue),
            Target_60: parseInt(ele.TargetValue60),
            Target_90: parseInt(ele.TargetValue90),
            Status: 'Y',
          });
        } else if (ele.TargetByModel_60 !== ele.TargetValue90) {
          finalarray.push({
            StoreID: ele.StoreID,
            ModelID: ele.ModelCode,
            StockType: stkTyp,
            Target: parseInt(ele.TargetValue),
            Target_60: parseInt(ele.TargetValue60),
            Target_90: parseInt(ele.TargetValue90),
            Status: 'Y',
          });
        }
      });
    });
    console.log(finalarray);
    this.shared.spinner.show();
    const obj = {
      targetwisemodel: finalarray,
    };
    if (finalarray.length > 0) {
      this.shared.api.postmethod('targetwisemodel/AddTargetWiseModelV2', obj).subscribe((res: any) => {
        if (res.status == 200) {
          this.shared.spinner.hide();
          // this.toast.success('Target Wise Model Updated Successfully');
          this.GetInventorySummaryReport();
          this.closeTarget()
        } else {
          // this.toast.success('Something went wrong please try again');
        }
      });
    } else {
      this.shared.spinner.hide();
      // this.toast.warning('No changes Found !');
    }
  }
  closeTarget() {
    // this.ngmodelactive.dismiss()
  }
  CompleteComponentState: boolean = true;
  subdataindex: any = 0;
  openDetails(data2: any, data3: any, store: any, block: any, newAge: any, usedAge: any,) {
    this.MainGrid = 'N'
    this.Block = block;
    this.var3 = data3;
    this.var2 = data2;
    this.ageCategory = this.stockType[0] == 'New' ? newAge : usedAge
    this.getDetails()
  }
  backtoWR() {
    this.MainGrid = 'Y'
  }
  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    }
    return false;
  }
  subscription!: Subscription;
  ngAfterViewInit(): void {
    // //console.log();
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.shared.common.pageName == 'Inventory Book') {
       if (res.obj.storesData != undefined) {
          this.groupsArray = res.obj.storesData;
          this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
          this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_name : this.groupName = ''
          this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
          this.getStoresandGroupsValues()
        }
      }
    })
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res: { obj: { title: string; stateEmailPdf: boolean; Email: any; notes: any; from: any; }; }) => {
      if (this.email != undefined) {
        if (res.obj.title == 'Inventory Book') {
          if (res.obj.stateEmailPdf == true) {
            // this.sendEmailData(res.obj.Email, res.obj.notes, res.obj.from);
          }
        }
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res: { obj: { title: string; statePrint: boolean; }; }) => {
      if (this.print != undefined) {
        if (res.obj.title == 'Inventory Book') {
          if (res.obj.statePrint == true) {
            // this.GetPrintData();
          }
        }
      }
    });
    this.Pdf = this.shared.api.getExportToPDFAllReports().subscribe((res: { obj: { title: string; statePDF: boolean; }; }) => {
      if (this.Pdf != undefined) {
        if (res.obj.title == 'Inventory Book') {
          if (res.obj.statePDF == true) {
            // this.generatePDF();
          }
        }
      }
    });
    this.excel = this.shared.api.getExportToExcelAllReports().subscribe((res: { obj: { title: string; state: boolean; }; }) => {
      if (this.excel != undefined) {
        if (res.obj.title == 'Inventory Book') {
          if (res.obj.state == true) {
            if (this.MainGrid == 'Y') {
              this.exportToExcel();
            } else {
              this.exportAsXLSX()
            }
          }
        }
      }
    });
  }
  ngOnDestroy() {
    this.unSubscribeing()
  }
  unSubscribeing() {
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
  getBackgroundColor(value: number): string {
    if (this.status.indexOf('Stock') >= 0 || this.status.indexOf('Stock_Transit') >= 0) {
      if (value < -3 || value > 3) {
        return '#f4b084';
      } else {
        return '#bdd7ee';
      }
    }
    return '#fff'
  }
  getUsedBackgroundColor(value: any): any {
    if (value == '[0-30]') {
      return '#f8cbad';
    } else if (value == '[31-45]') {
      return '#ddebf7';
    } else if (value == '[46-60]') {
      return '#c6e0b4';
    } else if (value == '[60+]') {
      return '#ed7d31';
    }
  }
  Scrollpercent: any = 0;
  scrollCurrentposition: any = 0;
  @ViewChild('scrollcent') scrollcent!: ElementRef;
  updateVerticalScroll(event: any): void {
    this.scrollCurrentposition = event.target.scrollTop;
    const scrollDemo = document.querySelector('#scrollcent') as HTMLElement;
    this.Scrollpercent = Math.round(
      (event.target.scrollTop /
        (event.target.scrollHeight - scrollDemo.clientHeight)) *
      100
    );
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
  stockType: any = ['New'];
  aged: any = ['MNM'];
  wholesale: any = ['N'];
  status: any = ['Stock'];
  multipleorsingle(block: any, e: any) {
    if (block == 'ST') {
      const index = this.stockType.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.stockType.splice(index, 1);
      } else {
        this.stockType = []
        // if(e=='New'){
        // this.aged = []
        // this.aged.push('MNM');
        // }
        // // if(e=='Used'){
        // //   this.aged = []
        // //   this.aged.push('AUY');
        // // }
        if (e == 'Fleet') {
          this.aged = []
          this.aged.push('MNM');
        }
        this.stockType.push(e);
      }
    }
    if (block == 'A') {
      const index = this.aged.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.aged.splice(index, 1);
      } else {
        this.aged = []
        // if(e=='MNM'){
        //   this.stockType=[]
        //   this.stockType.push('New')
        // }
        // if(e=='ANM'){
        //   this.stockType=[]
        //   this.stockType.push('New')
        // }
        // if(e=='AUY'){
        //   this.stockType=[]
        //   this.stockType.push('Used')
        // }
        this.aged.push(e);
      }
    }
    if (block == 'W') {
      this.wholesale = [];
      this.wholesale.push(e);
    }
    if (block == 'S') {
      const index = this.status.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.status.splice(index, 1);
      } else {
        this.status.push(e);
      }
    }
  }
  activePopover: number = -1;
  togglePopover(popoverIndex: number) {
    if (this.activePopover === popoverIndex) {
      // If the same popover is clicked, close it
      this.activePopover = -1;
    } else {
      // Open the selected popover and close others
      this.activePopover = popoverIndex;
    }
  }
  viewreport() {
    this.activePopover = -1
    if (this.stockType.length == 0) {
      alert('Please select atleast one stocktype');
    }
    if (this.status.length == 0) {
      alert('Please select atleast one status');
    }
    if (this.aged.length == 0) {
      alert('Please select atleast one additional');
    }
    else {
      this.setHeaderData()
      this.GetInventorySummaryReport();
    }
  }
  notesViewState: boolean = true
  notesView() {
    this.notesViewState = !this.notesViewState
  }
  notesData: any = {}
  popup: any
  addNotes(data: any, ref: any) {
    // this.notesData = {
    //   // store: this.inventorydetails[0].storeid,
    //   // mainkey: data.StockNo,
    //   // module: 'IN'
    //   module: '18',
    //   store: this.inventorydetails[0].storeid,
    //   title1: data.StockNo,
    //   title2: '',
    //   apiRoute: 'AddGeneralNotes'
    // }
    // this.popup = this.ngbmodal.open(ref, { size: 'xxl', backdrop: 'static' });
  }
  toggleView(data: any) {
    if (data.notesView == '+') {
      data.notesView = '-'
    } else {
      data.notesView = '+'
    }
  }
  closeNotes(e: any) {
    // this.ngbmodalActive.dismiss()
    this.popup.close()
    if (e == 'S') {
      this.getDetails()
    }
  }
  getDetails() {
    this.shared.spinner.show()
    const obj = {
      "As_id": this.storeIds,
      "StockType": this.stockType.toString(),
      "Status": this.status.toString(),
      "WholeSale": this.wholesale.toString(),
      "EXP": this.var3,
      "AgeCategory": this.ageCategory,
      'UserID': JSON.parse(localStorage.getItem('userInfo')!)?.user_Info?.userid,
    };
    this.shared.api
      .postmethod(this.shared.common.routeEndpoint + 'GetInventoryBookDetailsV1', obj)
      .subscribe((res) => {
        if (res.status == 200) {
          if (res.response && res.response.length > 0) {
            this.inventoryBookDetails = res.response.map((v: any) => ({
              ...v, notesView: '+'
            }));
            this.inventoryBookDetails.some(function (x: any) {
              if (x.Notes != undefined && x.Notes != null) {
                x.Notes = JSON.parse(x.Notes);
              }
            });
            this.shared.spinner.hide()
            this.DetailsNoData = ''
          } else {
            this.shared.spinner.hide()
            this.DetailsNoData = 'No Data Found!!'
          }
        } else {
          this.shared.spinner.hide()
          this.DetailsNoData = 'No Data Found!!'
        }
      });
  }
  exportToExcel() {
    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('Inventory Book');
    const titleRow = worksheet.addRow(['Inventory Book Report']);
    titleRow.font = { name: 'Arial', size: 14, bold: true };
    titleRow.alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.mergeCells('A1:N1');
    worksheet.addRow([]);
    let storeValue = 'All Stores';
    if (
      this.storeIds &&
      this.storeIds.length > 0 &&
      this.storeIds.length !== this.stores.length
    ) {
      storeValue = this.stores
        .filter((s: any) => this.storeIds.includes(s.ID))
        .map((s: any) => s.storename)
        .join(', ');
    }
    const filters = [
      { name: 'Store:', values: storeValue },
      { name: 'Stock Type:', values: this.stockType.toString() },
      { name: 'Units:', values: this.status.toString() },
      { name: 'Wholesale:', values: this.wholesale.map((item: any) => item === 'Y' ? 'Yes' : 'No').toString() },
      { name: 'View:', values: this.aged.toString() == 'AUY' ? 'AGED BY YEAR' : (this.aged.toString() == 'MNM' ? 'MAKE BY MODEL' : (this.aged.toString() == 'ANM' ? 'AGED BY MODEL' : '')) }
    ];
    let startIndex = 3;
    filters.forEach(f => {
      startIndex++;
      worksheet.addRow('');
      worksheet.getCell(`A${startIndex}`).value = f.name;
      worksheet.getCell(`B${startIndex}`).value = f.values;
      worksheet.getCell(`A${startIndex}`).font = { bold: true };
      worksheet.mergeCells(`B${startIndex}:G${startIndex}`);
    });
    startIndex++
    worksheet.addRow([]);
    startIndex++
    worksheet.addRow('');
    worksheet.getCell(`A${startIndex}`);
    worksheet.getCell(`B${startIndex}`);
    worksheet.getCell(`C${startIndex}`);
    worksheet.getCell(`D${startIndex}`);
    worksheet.getCell(`E${startIndex}`);
    worksheet.mergeCells(`F${startIndex}:G${startIndex}`);
    worksheet.getCell(`A${startIndex}`).value = 'New';
    worksheet.getCell(`B${startIndex}`).value = 'In Stock';
    worksheet.getCell(`C${startIndex}`).value = 'Target';
    worksheet.getCell(`D${startIndex}`).value = 'Over/Under';
    worksheet.getCell(`E${startIndex}`).value = '3-Mo Avg';
    worksheet.getCell(`F${startIndex}`).value = 'Age (In Days)';
    worksheet.getRow(1).height = 25;
    [`A${startIndex}`, `B${startIndex}`, `C${startIndex}`, `D${startIndex}`, `E${startIndex}`, `F${startIndex}`].forEach(key => {
      const cell = worksheet.getCell(key);
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2F5597' } };
    });
    const mainHeaders = ['', '', '', '', 'Sales', 'Oldest Units', 'Newest Units'];
    const mainhead = worksheet.addRow(mainHeaders);
    mainhead.font = { bold: true };
    mainhead.alignment = { horizontal: 'center', vertical: 'middle' };
    mainhead.eachCell((cell, number) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'dbe7ff' },
        bgColor: { argb: 'FF0000FF' },
      };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });
    const bindingHeaders = ['data3', 'InStock', 'Target', 'OUValue', 'AvgMonth', 'Oldest', 'Newest'];
    const currencyFields: any = [];
    for (const info of this.InventorySummaryData) {
      if (info.data2 != undefined) {
        for (const data2 of info.data2) {
          const nestedRowData = [data2.data2, '', '', '', '', '', '']
          const mainlayer = worksheet.addRow(nestedRowData);
          mainlayer.eachCell((cell, number) => {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'dbe7ff' },
              bgColor: { argb: 'FF0000FF' },
            };
          });
          if (data2.SubData != undefined) {
            for (const data3 of data2.SubData) {
              const nestedRowData = bindingHeaders.map(key => {
                const val = data3[key];
                return val === 0 || val == null ? '-' : '      ' + val;
              });
              const nestedRow = worksheet.addRow(nestedRowData);
              bindingHeaders.forEach((key, index) => {
                const cell = nestedRow.getCell(index + 1);
                if (currencyFields.includes(key) && typeof cell.value === 'number') {
                  cell.numFmt = '"$"#,##0.00';
                  cell.alignment = { horizontal: 'right' };
                } else if (!isNaN(Number(cell.value))) {
                  cell.alignment = { horizontal: 'right' };
                }
              });
            }
          }
        }
      }
    }
    const Data4 = worksheet.addRow([
      'Totals',
      this.getTotal('InStock'),
      this.getTotal('Target'),
      '',
      '',
      '',
      '',
    ]);
    Data4.outlineLevel = 1; // Grouping level 2
    Data4.font = { name: 'Arial', family: 4, size: 9 };
    Data4.alignment = { vertical: 'middle', horizontal: 'right' };
    Data4.getCell(1).alignment = {
      indent: 2,
      vertical: 'middle',
      horizontal: 'left',
    };
    Data4.eachCell((cell: any, number: any) => {
      cell.border = { right: { style: 'dotted' } };
    });
    Data4.eachCell((cell: any, number: any) => {
      cell.border = { right: { style: 'dotted' } };
      if (number == 5) {
        cell.numFmt = '#,##0.0';
      }
    });
    if (Data4.number % 2) {
      Data4.eachCell((cell, number) => {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'e5e5e5' },
          bgColor: { argb: 'FF0000FF' },
        };
      });
    }

    worksheet.columns.forEach((col: any, i: number) => {
      col.width = i === 0 ? 30 : 15;
    });
    workbook.xlsx.writeBuffer().then((buffer: any) => {
      const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      this.shared.exportToExcel(workbook, 'Inventory Book');
    });
  }
  exportAsXLSX() {
    let localarray = this.inventoryBookDetails.map((_arrayElement: any) =>
      Object.assign({}, _arrayElement)
    );
    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('Inventory Book Details');
    let storeValue = 'All Stores';
    if (
      this.storeIds &&
      this.storeIds.length > 0 &&
      this.storeIds.length !== this.stores.length
    ) {
      storeValue = this.stores
        .filter((s: any) => this.storeIds.includes(s.ID))
        .map((s: any) => s.storename)
        .join(', ');
    }
    const filters = [
      { name: 'Store:', values:storeValue},
      { name: 'Stock Type:', values: this.stockType.toString() },
      { name: 'Units:', values: this.status.toString() },
      { name: 'Wholesale:', values: this.wholesale.map((item: any) => item === 'Y' ? 'Yes' : 'No').toString() },
      { name: 'View:', values: this.aged.toString() == 'AUY' ? 'AGED BY YEAR' : (this.aged.toString() == 'MNM' ? 'MAKE BY MODEL' : (this.aged.toString() == 'ANM' ? 'AGED BY MODEL' : '')) },
      { name: 'Model:', values: this.var3 },
    ];
    const ReportFilter = worksheet.addRow(['Inventory Book Details :']);
    ReportFilter.font = { name: 'Arial', family: 4, size: 10, bold: true };
    let startIndex = 2
    filters.forEach((val: any) => {
      startIndex++
      worksheet.addRow('');
      worksheet.getCell(`A${startIndex}`);
      worksheet.mergeCells(`B${startIndex}:C${startIndex}`);
      worksheet.getCell(`A${startIndex}`).value = val.name;
      worksheet.getCell(`B${startIndex}`).value = val.values
    })
    startIndex++
    worksheet.addRow('');
    startIndex++
    // const secondHeader = ['Stock #', 'Year', 'Make', 'Model', 'Color', 'Age', 'MSRP', 'VIN'];
    // worksheet.addRow(secondHeader);
    // worksheet.getRow(startIndex).height = 25;
    // worksheet.getRow(startIndex).font = { bold: true, color: { argb: 'FFFFFFFF' } };
    // worksheet.getRow(startIndex).alignment = { vertical: 'middle', horizontal: 'center' };
    // worksheet.getRow(startIndex).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2F5597' } };
    const secondHeader = [
      'Stock #', 'Year', 'Make', 'Model',
      'Color', 'Age', 'MSRP', 'VIN'
    ];
    
    const headerRow = worksheet.addRow(secondHeader);
    headerRow.height = 25;
    
    headerRow.eachCell((cell) => {
      cell.font = {
        bold: true,
        color: { argb: 'FFFFFFFF' }
      };
    
      cell.alignment = {
        vertical: 'middle',
        horizontal: 'center'
      };
    
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF2F5597' }
      };
    });
    
    const bindingHeaders = ['StockNo', 'Year', 'Make', 'Model', 'Color', 'Age', 'MSRP', 'VIN'];
    const currencyFields: any = ['MSRP'];
    let notesCount = 12
    for (const info of localarray) {
      const rowData = bindingHeaders.map(key => {
        const val = info[key];
        return (val === 0 || val == null || val == '' ? '-' : val);
      });
      const dealerRow = worksheet.addRow(rowData);
      // dealerRow.font = { bold: false };
      bindingHeaders.forEach((key, index) => {
        const cell = dealerRow.getCell(index + 1);
        if (currencyFields.includes(key) && typeof cell.value === 'number') {
          cell.numFmt = '"$"#,##0.00';
          cell.alignment = { horizontal: 'right' };
        } else if (!isNaN(Number(cell.value))) {
          cell.alignment = { horizontal: 'center' };
        } else {
          cell.alignment = { horizontal: 'center' };
        }
      });
      if (info.NotesStatus == 'Y' && this.notesViewState == true) {
        worksheet.mergeCells(notesCount, 1, notesCount, 8);
        const Data2NOtes = worksheet.getCell(notesCount, 1);
        Data2NOtes.value = 'Notes'
        Data2NOtes.alignment = { indent: 2, vertical: 'middle', horizontal: 'left', };
        Data2NOtes.font = { name: 'Arial', family: 4, size: 9 };
        Data2NOtes.border = { right: { style: 'thin' }, left: { style: 'thin' }, top: { style: 'thin' }, bottom: { style: 'thin' } };
        notesCount++
        for (const d1 of info.Notes) {
          worksheet.mergeCells(notesCount, 1, notesCount, 8);
          const Data2 = worksheet.getCell(notesCount, 1);
          Data2.value = ' ' + ' ' + ' ' + ' ' + ' ' + ' ' + d1.N_Notes
          Data2.alignment = { indent: 2, vertical: 'middle', horizontal: 'left', };
          Data2.font = { name: 'Arial', family: 4, size: 9 };
          Data2.border = { right: { style: 'thin' }, left: { style: 'thin' }, top: { style: 'thin' }, bottom: { style: 'thin' } };
          Data2.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'b6d3ec' },
            bgColor: { argb: 'b4c7dc' },
          };
          notesCount++
        }
      }
    }


    worksheet.columns.forEach((column: any) => {
      let maxLength = 25;
      column.width = maxLength + 2;
    });
    workbook.xlsx.writeBuffer().then((buffer: any) => {
      const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      this.shared.exportToExcel(workbook, 'Inventory Book Details')
    });
  }
}
