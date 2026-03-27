import { Component, OnInit, ViewChild, ElementRef, HostListener } from '@angular/core';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { Setdates } from '../../../../Core/Providers/SetDates/setdates';
import { common } from '../../../../common';
import { Subscription } from 'rxjs';
import { environment } from '../../../../../environments/environment';
import { Stores } from '../../../../CommonFilters/stores/stores';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
import { CurrencyPipe } from '@angular/common';
import { Workbook } from 'exceljs';

@Component({
  selector: 'app-dashboard',
  imports: [SharedModule, Stores],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss'
})
export class Dashboard {

  PartsData: any = [];
  IndividualPartsGross: any = [];
  TotalPartsGross: any = [];
  PartsSource: any = []
  selectedpartssource: any = [];
  TotalReport: any = 'B';
  NoData: boolean = false;
  DateType: any = '30';
  responcestatus: any = '';
  partsDetails: any = 'Y'

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


  Pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;
  // index: any;
  otherstoreid: any = '';
  selectedotherstoreids: any = '';
  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent) {
    const clickedInside = (event.target as HTMLElement).closest('.dropdown-toggle, .reportstores-card, .timeframe');
    if (!clickedInside) {
      this.activePopover = -1;
    }
  }

  constructor(
    public shared: Sharedservice, public setdates: Setdates, private comm: common, private cp: CurrencyPipe, private toast: ToastService,
  ) {
    this.shared.setTitle(this.comm.titleName + '-Parts Aging');
    if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
      this.groupId = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Preferences
      this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids.split(',')
    }
    if (this.shared.common.groupsandstores.length > 0) {
      this.groupsArray = this.shared.common.groupsandstores.filter((val: any) => val.sg_id != this.shared.common.reconID);
      this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
      this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_Name : this.groupName = ''
      this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
      this.getStoresandGroupsValues()
    }
    this.setHeaderData()
    this.getPartssource()
    this.GetData();

  }

  ngOnInit(): void {
    // var curl = 'https://fbxtract.axelautomotive.com/favouritereports/GetPartsAgingReportV1';
    // this.shared.api.logSaving(curl,{},'','Success','Parts Aging');
  }

  setHeaderData() {
    const data = {
      title: 'Parts Aging',
      stores: this.storeIds,
      partssource: this.PartsSource,
      ToporBottom: this.TotalReport,
      groups: this.groupId,
    };
    this.shared.api.SetHeaderData({
      obj: data,
    });
  }
  getPartssource(count?: any) {
    const obj = {
      StoreID: this.storeIds.toString()
    };
    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetPartsSource', obj).subscribe(
      (res) => {
        if (res.status == 200) {
          this.PartsSource = res.response.filter((x: any) => x.AP_Source != '')
          this.selectedpartssource = this.PartsSource.map(function (a: any) {
            return a.AP_Source;
          });
        } else {
          this.toast.show('Invalid Details','danger','Error');
        }
      },
      (error) => {
        // //console.log(error);
      }
    );
  }

  allsource() {
    if (this.selectedpartssource.length == this.PartsSource.length) {
      this.selectedpartssource = []
    } else {
      this.selectedpartssource = this.PartsSource.map(function (a: any) {
        return a.AP_Source;
      });
    }
  }
  getPartsData() {
    this.NoData = false
    this.responcestatus = '';
    this.IndividualPartsGross = [];
    this.TotalPartsGross = [];
    this.shared.spinner.show();
    if (this.storeIds != '' || this.selectedotherstoreids != '') {
      this.GetData();
    } else {
      this.NoData = true;
      this.shared.spinner.hide()
    }
    // this.GetTotalData();
  }

  GetData() {
    this.IndividualPartsGross = [];
    this.shared.spinner.show();
    const obj = {
      StoreID: this.selectedotherstoreids != undefined && this.selectedotherstoreids != '' && this.selectedotherstoreids != null ?
        (this.storeIds != '' ? this.storeIds + ',' + this.selectedotherstoreids.toString() : this.selectedotherstoreids.toString()) : this.storeIds,
      UserID: 0,
      Source: this.selectedpartssource.toString(),
      DateType: this.DateType
    };
    const curl = environment.apiUrl + this.comm.routeEndpoint + 'GetPartsAgingReportV1';
    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetPartsAgingReportV1', obj).subscribe(
      (res) => {
        const currentTitle = document.title;
        this.shared.api.logSaving(curl, {}, '', res.message, currentTitle);
        if (res.status == 200) {
          if (res.response != undefined) {
            if (res.response.length > 0) {
              this.shared.spinner.hide();
              this.IndividualPartsGross = [];
              this.IndividualPartsGross = res.response;
              //console.log(this.IndividualPartsGross)
              this.responcestatus = this.responcestatus + 'I';
              let idi_len = this.IndividualPartsGross.length;
              this.IndividualPartsGross.some(function (x: any) {
                if (x.AgeGroupData != undefined && x.AgeGroupData != '') {
                  x.AgeGroupData = JSON.parse(x.AgeGroupData);
                  x.Dealer = '-';
                }
                // if (idi_len == 1) {
                //   x.Dealer = '-';
                // } else {
                //   x.Dealer = '+';
                // }
              });
              console.log(this.IndividualPartsGross, 'chkkk')

              // this.combineIndividualandTotal();
              // if (this.TotalReport == 'T') {
              //   let last = this.IndividualPartsGross.pop()
              //   this.IndividualPartsGross.unshift(last)
              // }

              const idx = this.IndividualPartsGross.findIndex(
                (x: any) => x?.StoreName === 'REPORT TOTAL'
              );

              if (idx > -1) {
                const [reportTotalRow] = this.IndividualPartsGross.splice(idx, 1); // remove it

                if (this.TotalReport === 'T') {
                  this.IndividualPartsGross.unshift(reportTotalRow); // bring to top
                } else {
                  this.IndividualPartsGross.push(reportTotalRow);    // bring to bottom
                }
              }
              //               else {
              //   const first = this.IndividualPartsGross.shift();
              //   if (first !== undefined) this.IndividualPartsGross.push(first);
              // }
            }
            else {
              // this.toast.error('Empty Response','');
              this.shared.spinner.hide();
              this.NoData = true;

            }
          } else {
            // this.toast.error('Empty Response');
            this.shared.spinner.hide();
            this.NoData = true;
          }
        } else {
          this.toast.show(res.status, 'danger', 'Error');
          this.shared.spinner.hide();
          this.NoData = true;

        }
      },
      (error) => {
        this.toast.show('502 Bad Gate Way Error', 'danger', 'Error');
        this.shared.spinner.hide();
        this.NoData = true;
      }
    );
  }




  calculateQtyPer(qty: number, total: number): number {
    if (!qty || !total || total === 0) {
      return 0;
    }
    return (qty * 100) / total;
  }

  calculatestock(saleinfo: any) {
    if (!saleinfo || !saleinfo.AgeGroupData || saleinfo.Stock_Cost == 0) return 0;
    const ageGroup12Plus = saleinfo.AgeGroupData.find((a: any) => a.Never_sold === '12+');
    if (ageGroup12Plus && ageGroup12Plus.Stock_Cost != null) {
      return (ageGroup12Plus.Stock_Cost * 100) / saleinfo.Stock_Cost;
    }
    return 0;
  }
  calculatenonstock(saleinfo: any) {
    if (!saleinfo || !saleinfo.AgeGroupData || saleinfo.NonStock_Cost == 0) return 0;
    const ageGroup12Plus = saleinfo.AgeGroupData.find((a: any) => a.Never_sold === '12+');
    if (ageGroup12Plus && ageGroup12Plus.NonStock_Cost != null) {
      return (ageGroup12Plus.NonStock_Cost * 100) / saleinfo.NonStock_Cost;
    }
    return 0;
  }
  calculatetotal(saleinfo: any) {
    if (!saleinfo || !saleinfo.AgeGroupData || saleinfo.Total_Cost == 0) return 0;
    const ageGroup12Plus = saleinfo.AgeGroupData.find((a: any) => a.Never_sold === '12+');
    if (ageGroup12Plus && ageGroup12Plus.Total_Cost != null) {
      return (ageGroup12Plus.Total_Cost * 100) / saleinfo.Total_Cost;
    }
    return 0;
  }

  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    }
    else if (value < 0) {
      return false;
    }
    return true
  }
  subdataindex: any = 0;
  expandorcollapse(ind: any, e: any, ref: any, Item: any, parentData: any) {
    let id = (e.target as Element).id;
    if (id == 'D_' + ind) {

      if (ref == '-') {
        Item.Dealer = '+';
      }
      if (ref == '+') {
        Item.Dealer = '-';
      }

    }
    if (id == 'DN_' + ind) {
      if (ref == '-') {
        Item.data2sign = '+';
      }
      if (ref == '+') {
        Item.data2sign = '-';
        Item.Dealer = '-';
      }
    }
  }
  selectedData: any = { store_id: '', daterange: '' }
  daterange: any;

  openDetails(Item: any, ParentItem: any, ref: any) {
    if (ParentItem.StoreName != 'REPORT TOTAL') {
      if (ref == '2') {
        this.selectedData.daterange = Item.Never_sold
        this.selectedData.store_id = ParentItem.Store
          this.daterange = Item.Never_sold?.replace(' months', '');
        this.GetDetails()
      }
    }

  }
  details: any = []
  PartsPersonDetails: any = []
  spinnerLoader: boolean = false;
  spinnerLoadersec: boolean = false
  GetDetails() {
    this.partsDetails = 'N'
    this.PartsPersonDetails= []
    this.details= []
    this.shared.spinner.show()
    const obj = {
      store_id: this.selectedData.store_id,
      daterange: this.selectedData.daterange,
      // PageNumber: this.pageNumber,
      // PageSize: '100',

    };
    this.shared.api
      .postmethod(this.comm.routeEndpoint + 'GetFbPartsAgingJSONDetails', obj)
      .subscribe((res) => {
        if (res.status == 200) {
          this.details = res.response;
          this.PartsPersonDetails = [
            ...this.PartsPersonDetails,
            ...this.details,
          ];
          console.log(this.PartsPersonDetails);
          this.shared.spinner.hide()
          this.spinnerLoader = false;
          this.spinnerLoadersec = false;
          // this.PartsPersonDetails=res.response
          console.log(this.PartsPersonDetails);

          this.PartsPersonDetails.some(function (x: any) {
            if (x.PartsJsonDetails != undefined && x.PartsJsonDetails != '') {
              x.PartsJsonDetails = JSON.parse(x.PartsJsonDetails);
              x.Dealer = '-';
            }
          });


          // this.shared.spinner.hide()
          this.spinnerLoader = false;
          if (this.PartsPersonDetails.length > 0) {
            this.NoData = false;
          } else {
            this.NoData = true;
          }
        }
      });
  }
backtoWR(){
  this.partsDetails = 'Y'
}
  isDesc: boolean = false;
  column: string = 'CategoryName';

  sort(property: any) {
    this.isDesc = !this.isDesc; //change the direction
    this.column = property;
    let direction = this.isDesc ? 1 : -1;
    // //console.log(property)
    this.PartsData.sort(function (a: any, b: any) {
      if (a[property] < b[property]) {
        return -1 * direction;
      } else if (a[property] > b[property]) {
        return 1 * direction;
      } else {
        return 0;
      }
    });
  }

  ngAfterViewInit(): void {

    this.shared.api.getStores().subscribe((res: any) => {
      if (this.shared.common.pageName == 'Parts Aging') {
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
        if (res.obj.title == 'Parts Aging') {
          if (res.obj.state == true) {
            this.exportToExcel();
          }
        }
      }
    });

    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res) => {
      if (this.email != undefined) {
        if (res.obj.title == 'Parts Aging') {
          if (res.obj.stateEmailPdf == true) {
            //   this.sendEmailData(res.obj.Email, res.obj.notes, res.obj.from);
          }
        }
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res) => {
      if (this.print != undefined) {
        if (res.obj.title == 'Parts Aging') {
          if (res.obj.statePrint == true) {
            //    this.GetPrintData();
          }
        }
      }
    });
    this.Pdf = this.shared.api.getExportToPDFAllReports().subscribe((res) => {
      if (this.Pdf != undefined) {
        if (res.obj.title == 'Parts Aging') {
          if (res.obj.statePDF == true) {
            //  this.generatePDF();
          }
        }
      }
    });
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
  Selectedpartssource(e: any) {
    const index = this.selectedpartssource.findIndex((i: any) => i == e.AP_Source);
    if (index >= 0) {
      this.selectedpartssource.splice(index, 1);
    } else {
      this.selectedpartssource.push(e.AP_Source);
    }
  }

  StoresData(data: any) {
    this.storeIds = data.storeids;
    this.groupId = data.groupId;
    this.storename = data.storename;
    this.groupName = data.groupName;
    this.storecount = data.storecount;
    this.storedisplayname = data.storedisplayname;
    this.getPartssource()
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
    if (block == 'TB') {
      this.TotalReport = e;
    }
  }
  activePopover: number = -1;

  togglePopover(popoverIndex: number) {
    this.activePopover = this.activePopover === popoverIndex ? -1 : popoverIndex;
  }
  viewreport() {
    this.activePopover = -1

    if (this.storeIds.length == 0 && this.selectedotherstoreids.length == 0) {
      this.toast.show('Please select atleast one Store', 'warning', 'Warning');
    } else {
      this.setHeaderData();
      this.GetData()
    }

  }
  ExcelStoreNames: any = [];
  exportToExcel() {
    let storeNames: any = [];
    let store = this.storeIds.toString()
    // const obj = {
    //   id: this.groups,
    //   userid: JSON.parse(localStorage.getItem('UserDetails')!).userid,
    // };
    // this.shared.api
    //   .postmethodOne(this.comm.routeEndpoint+'GetStoresbyGroupuserid', obj)
    //   .subscribe((res: any) => {
    storeNames = this.comm.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores.filter((item: any) =>
      store.some((cat: any) => cat === item.ID.toString())
    );
    // //console.log(store,res.response.length);

    if (store.length == this.comm.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores.length) {
      this.ExcelStoreNames = 'All Stores'
    } else {
      this.ExcelStoreNames = storeNames.map(function (a: any) {
        return a.storename;
      });
    }
    const PartsData = this.IndividualPartsGross.map((_arrayElement: any) =>
      Object.assign({}, _arrayElement)
    );

    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('Parts Aging');
    worksheet.views = [
      {
        state: 'frozen',
        ySplit: 12, // Number of rows to freeze (2 means the first two rows are frozen)
        topLeftCell: 'A13', // Specify the cell to start freezing from (in this case, the third row)
        showGridLines: false,
      },
    ];
    worksheet.addRow('');
    const titleRow = worksheet.addRow(['Parts Aging']);
    titleRow.eachCell((cell: any, number: any) => {
      cell.alignment = { indent: 1, vertical: 'middle', horizontal: 'left' };
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
    const Groups = worksheet.getCell('B6');
    Groups.value = 'Groups :';
    const groups = worksheet.getCell('D6');
    groups.value =
      this.comm.groupsandstores.filter((val: any) => val.sg_id == this.groupId.toString())[0].sg_name;
    // this.groups == 1 ? this.comm.excelName : (this.groups == 2 ? 'Domestic' : (this.groups == 3 ? 'Import' : (this.groups == 4 ? 'Warehouse' : '-')));
    groups.font = { name: 'Arial', family: 4, size: 9 };
    const Brands = worksheet.getCell('B7');
    Brands.value = 'Brands :';
    const brands = worksheet.getCell('D7');
    brands.value = '-';
    brands.font = { name: 'Arial', family: 4, size: 9 };
    const Stores1 = worksheet.getCell('B8');
    Stores1.value = 'Stores :';
    worksheet.mergeCells('D8', 'O10');
    const stores1 = worksheet.getCell('D8');
    stores1.value =
      this.ExcelStoreNames == 0
        ? 'All Stores'
        : this.ExcelStoreNames == null
          ? '-'
          : this.ExcelStoreNames.toString().replaceAll(',', ', ');
    stores1.font = { name: 'Arial', family: 4, size: 9 };
    stores1.alignment = { vertical: 'top', horizontal: 'left', wrapText: true };

    worksheet.addRow('');

    let dateYear = worksheet.getCell('A11');
    dateYear.value = 'As of ';
    dateYear.alignment = { vertical: 'middle', horizontal: 'center' };
    dateYear.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
      color: { argb: 'FFFFFF' },
    };
    dateYear.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '2a91f0' },
      bgColor: { argb: 'FF0000FF' },
    };
    dateYear.border = { right: { style: 'thin' } };

    worksheet.mergeCells('B11', 'D11');
    let totalparts = worksheet.getCell('B11');
    totalparts.value = 'Stock';
    totalparts.alignment = { vertical: 'middle', horizontal: 'center' };
    totalparts.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
      color: { argb: 'FFFFFF' },
    };
    totalparts.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '2a91f0' },
      bgColor: { argb: 'FF0000FF' },
    };
    totalparts.border = { right: { style: 'thin' } };

    worksheet.mergeCells('E11', 'G11');
    let mechanical = worksheet.getCell('E11');
    mechanical.value = 'Non-Stock';
    mechanical.alignment = { vertical: 'middle', horizontal: 'center' };
    mechanical.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
      color: { argb: 'FFFFFF' },
    };
    mechanical.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '2a91f0' },
      bgColor: { argb: 'FF0000FF' },
    };
    mechanical.border = { right: { style: 'thin' } };

    worksheet.mergeCells('H11', 'J11');
    let retail = worksheet.getCell('H11');
    retail.value = 'Total';
    retail.alignment = { vertical: 'middle', horizontal: 'center' };
    retail.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
      color: { argb: 'FFFFFF' },
    };
    retail.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '2a91f0' },
      bgColor: { argb: 'FF0000FF' },
    };
    retail.border = { right: { style: 'thin' } };



    let Headings = [
      '',

      'Qty',
      'Cost',
      'Idle',

      'Qty',
      'Cost',
      'Idle',

      'Qty',
      'Cost',
      'Idle',

    ];
    const headerRow = worksheet.addRow(Headings);
    headerRow.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: false,
      color: { argb: 'FFFFFF' },
    };
    headerRow.alignment = { indent: 1, vertical: 'middle', horizontal: 'center' };
    headerRow.eachCell((cell: any, number: any) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '788494' },
        bgColor: { argb: 'FF0000FF' },
      };
      cell.border = { right: { style: 'thin' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });

    for (const d of PartsData) {
      const Data1 = worksheet.addRow([
        d.StoreName == '' ? '-' : d.StoreName == null ? '-' : d.StoreName,

        d.Stock_Qty == '' ? '-' : d.Stock_Qty == null ? '-' : d.Stock_Qty,
        d.Stock_Cost == ''
          ? '-'
          : d.Stock_Cost == null
            ? '-'
            : d.Stock_Cost,
        d.Stock_Idle == ''
          ? '-'
          : d.Stock_Idle == null
            ? '-'
            : d.Stock_Idle,

        d.NonStock_Qty == '' ? '-' : d.NonStock_Qty == null ? '-' : d.NonStock_Qty,
        d.NonStock_Cost == '' ? '-' : d.NonStock_Cost == null ? '-' : d.NonStock_Cost,
        d.NonStock_Idle == '' ? '-' : d.NonStock_Idle == null ? '-' : d.NonStock_Idle,

        d.Total_Qty == '' ? '-' : d.Total_Qty == null ? '-' : d.Total_Qty,
        d.Total_Cost == ''
          ? '-'
          : d.Total_Cost == null
            ? '-'
            : d.Total_Cost,
        d.Total_Idle == ''
          ? '-'
          : d.Total_Idle == null
            ? '-'
            : d.Total_Idle,
      ]);
      // Data1.outlineLevel = 1; // Grouping level 1
      Data1.font = { name: 'Arial', family: 4, size: 9 };
      Data1.getCell(1).alignment = {
        indent: 1,
        vertical: 'middle',
        horizontal: 'left',
      };
      Data1.eachCell((cell: any, number: any) => {
        cell.border = { right: { style: 'thin' } };
        if (number == 2 || number == 5 || number == 8) {
          cell.numFmt = '#,##0';
          cell.alignment = { vertical: 'middle', horizontal: 'center', indent: 1 };
        } if (number > 2 && number < 5) {
          cell.numFmt = '$#,##0';
          cell.alignment = { vertical: 'middle', horizontal: 'center', indent: 1 };
        } if (number > 5 && number < 8) {
          cell.numFmt = '$#,##0';
          cell.alignment = { vertical: 'middle', horizontal: 'center', indent: 1 };
        } if (number > 8 && number < 11) {
          cell.numFmt = '$#,##0';
          cell.alignment = { vertical: 'middle', horizontal: 'center', indent: 1 };
        }
        if (number != 1) {
          cell.alignment = { vertical: 'middle', horizontal: 'center', indent: 1 };
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
      if (d.Total_PartsSale < 0) {
        Data1.getCell(2).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
      } if (d.Total_PartsGross < 0) {
        Data1.getCell(4).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
      } if (d.Total_PartsGross_Pace < 0) {
        Data1.getCell(5).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
      } if (d.Total_PartsGrossTarget < 0) {
        Data1.getCell(6).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
      } if (d.Total_PartsGross_Diff < 0) {
        Data1.getCell(7).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
      } if (d.ServiceGross < 0) {
        Data1.getCell(8).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
      } if (d.ServiceGross_Pace < 0) {
        Data1.getCell(9).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
      } if (d.ServiceGross_Target < 0) {
        Data1.getCell(10).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
      } if (d.Total_PartsGross_Diff < 0) {
        Data1.getCell(11).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
      } if (d.PartsGross < 0) {
        Data1.getCell(12).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
      } if (d.PartsGross_Pace < 0) {
        Data1.getCell(13).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
      } if (d.PartsGross_Target < 0) {
        Data1.getCell(14).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
      } if (d.PartsGross_Diff < 0) {
        Data1.getCell(15).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
      } if (d.Parts_RO < 0) {
        Data1.getCell(16).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
      } if (d.Lost_PerDay < 0) {
        Data1.getCell(17).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
      } if (d.Retention < 0) {
        Data1.getCell(18).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
      }
      if (d.AgeGroupData != undefined) {
        for (const d1 of d.AgeGroupData) {
          const Data2 = worksheet.addRow([
            d1.Never_sold == '' ? '-' : d1.Never_sold == null ? '-' : d1.Never_sold,

            d1.Stock_Qty == '' ? '-' : d1.Stock_Qty == null ? '-' : d1.Stock_Qty,
            d1.Stock_Cost == ''
              ? '-'
              : d1.Stock_Cost == null
                ? '-'
                : d1.Stock_Cost,
            d1.Stock_Idle == ''
              ? '-'
              : d1.Stock_Idle == null
                ? '-'
                : d1.Stock_Idle,

            d1.NonStock_Qty == '' ? '-' : d1.NonStock_Qty == null ? '-' : d1.NonStock_Qty,
            d1.NonStock_Cost == '' ? '-' : d1.NonStock_Cost == null ? '-' : d1.NonStock_Cost,
            d1.NonStock_Idle == '' ? '-' : d1.NonStock_Idle == null ? '-' : d1.NonStock_Idle,

            d1.Total_Qty == '' ? '-' : d1.Total_Qty == null ? '-' : d1.Total_Qty,
            d1.Total_Cost == ''
              ? '-'
              : d1.Total_Cost == null
                ? '-'
                : d1.Total_Cost,
            d1.Total_Idle == ''
              ? '-'
              : d1.Total_Idle == null
                ? '-'
                : d1.Total_Idle,

          ]);
          Data2.outlineLevel = 1; // Grouping level 2
          Data2.font = { name: 'Arial', family: 4, size: 8 };
          Data2.getCell(1).alignment = {
            indent: 2,
            vertical: 'middle',
            horizontal: 'left',
          };
          Data2.eachCell((cell: any, number: any) => {
            cell.border = { right: { style: 'thin' } };
            if (number == 2 || number == 5 || number == 8) {
              cell.numFmt = '#,##0';
              cell.alignment = { vertical: 'middle', horizontal: 'center', indent: 1 };
            } if (number > 2 && number < 5) {
              cell.numFmt = '$#,##0';
              cell.alignment = { vertical: 'middle', horizontal: 'center', indent: 1 };
            } if (number > 5 && number < 8) {
              cell.numFmt = '$#,##0';
              cell.alignment = { vertical: 'middle', horizontal: 'center', indent: 1 };
            } if (number > 8 && number < 11) {
              cell.numFmt = '$#,##0';
              cell.alignment = { vertical: 'middle', horizontal: 'center', indent: 1 };
            }
            if (number != 1) {
              cell.alignment = {
                vertical: 'middle',
                horizontal: 'center',
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
          if (d1.Total_PartsSale < 0) {
            Data2.getCell(2).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
          } if (d1.Total_PartsGross < 0) {
            Data2.getCell(4).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
          } if (d1.Total_PartsGross_Pace < 0) {
            Data2.getCell(5).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
          } if (d1.Total_PartsGrossTarget < 0) {
            Data2.getCell(6).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
          } if (d1.Total_PartsGross_Diff < 0) {
            Data2.getCell(7).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
          } if (d1.ServiceGross < 0) {
            Data2.getCell(8).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
          } if (d1.ServiceGross_Pace < 0) {
            Data2.getCell(9).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
          } if (d1.ServiceGross_Target < 0) {
            Data2.getCell(10).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
          } if (d1.Total_PartsGross_Diff < 0) {
            Data2.getCell(11).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
          } if (d1.PartsGross < 0) {
            Data2.getCell(12).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
          } if (d1.PartsGross_Pace < 0) {
            Data2.getCell(13).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
          } if (d1.PartsGross_Target < 0) {
            Data2.getCell(14).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
          } if (d1.PartsGross_Diff < 0) {
            Data2.getCell(15).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
          } if (d1.Parts_RO < 0) {
            Data2.getCell(16).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
          } if (d1.Lost_PerDay < 0) {
            Data2.getCell(17).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
          } if (d1.Retention < 0) {
            Data2.getCell(18).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
          }
        }
      }
      if (d.data1 === 'Reports Total') {
        Data1.eachCell((cell) => {
          cell.font = { name: 'Arial', family: 4, size: 9, bold: true };
          // cell.border = {
          //   top: { style: 'thin' },
          //   bottom: { style: 'thin' },
          // };
        });
      }
    }

    worksheet.eachRow((row, rowIndex) => {
      row.eachCell((cell, colIndex) => {
        if (rowIndex > 1 && rowIndex < 19) { // Skip the header row
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
    worksheet.addRow([]);

    this.shared.exportToExcel(workbook, 'Parts Aging_' + DATE_EXTENSION);

  }


}


