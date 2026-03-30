import { Component, OnInit, ViewChild, ElementRef, HostListener } from '@angular/core';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { Setdates } from '../../../../Core/Providers/SetDates/setdates';
import { BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { common } from '../../../../common';
// import { Stores } from '../../../../CommonFilters/stores/stores';

import { Subscription } from 'rxjs';
import { Stores } from '../../../../CommonFilters/stores/stores';
import { Router } from '@angular/router'
import { group, log } from 'console';
import { CurrencyPipe } from '@angular/common';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [SharedModule, Stores],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss'
})
export class Dashboard {
  storeIds: any = '';
  TotalReport: any = 'T';
  date: any;
  AgeFrom: any = 0;
  AgeTo: any = 0;
  NoData: any;
  groups: any = 1;
  MainGrid: any = 'Y'

  path1name: any = 'Dealership';
  path2name: any = '';
  path3name: any = '';

  dealType: any = ['New'];
  stocktype: any = ["All"]
  header: any = [{ type: 'Bar', storeIds: this.storeIds, groups: this.groups }]
  statustype: any = ['Stock']
  ZeroBalance: any = ['N']
  Wholesale: any = ['N']

  popup: any = [{ type: 'Popup' }];
  Performance: any = 'Load';
  reportOpen!: Subscription;
  reportGetting!: Subscription;
  Pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;
  storename: any = '';
  range: any;
  totalcount: any = ''
  groupsArray: any = [];
  // storename: any = ''
  
  stores: any = [];
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  groupId: any = 0;

  storesFilterData: any = {
    'groupsArray': this.groupsArray, 'groupId': this.groupId, 'storesArray': this.stores, 'storeids': '1', 'type': 'M', 'others': 'N',
    'groupName': this.groupName, 'storename': this.storename, storecount: null, 'storedisplayname': this.storedisplayname
  };
  constructor(public shared: Sharedservice, public setdates: Setdates, private router: Router, private cp: CurrencyPipe,private toast: ToastService,) {
    if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
      // this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.ustores.split(',')
      // this.storeIds = '1,2'
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
    this.date = new Date();


    this.shared.setTitle(this.shared.common.titleName + "-Inventory Summary");
    if (typeof window !== 'undefined') {
      if (localStorage.getItem('DetailsObject') != null && localStorage.getItem('DetailsObject') != undefined) {
        const InvObj = JSON.parse(localStorage.getItem('DetailsObject')!);
        if (InvObj.dataobj.Data2 == 'AX1INS') {
          this.storeIds = InvObj.dataobj.Data1
          this.dealType = []
          this.dealType.push(InvObj.dataobj.val == 1 || InvObj.dataobj.val == 4 || InvObj.dataobj.val == 5 ? 'New' : (InvObj.dataobj.val == 2 ? 'Used' : 'Fleet'))
          this.stocktype = []
          this.stocktype = InvObj.dataobj.val == 4 ? 'RETAIL' : (InvObj.dataobj.val == 5 ? 'COMMERCIAL TRUCK' : 'All')
          this.Wholesale = ['N']
        }
      }


      
    }
    const data = {
      title: 'Inventory Summary',
      stores: this.storeIds,
      toporbottom: this.TotalReport,
      AgeFrom: this.AgeFrom,
      AgeTo: this.AgeTo,
      Month: this.date,
      groups: this.groups,
      dealType: this.dealType.toString(),
      stocktype: this.stocktype,
      statustype: this.statustype, ZeroBalance: this.ZeroBalance.toString(), Wholesale: this.Wholesale,
      count: 0

    };

    this.shared.api.SetHeaderData({ obj: data });
    this.header = [{
      type: 'Bar', storeIds: this.storeIds, groups: this.groups, statustype: this.statustype, ZeroBalance: this.ZeroBalance.toString(), Wholesale: this.Wholesale,
    }]
    this.getServiceData();

  }
  isDesc: boolean = false;
  column: string = 'CategoryName';
  ngOnInit(): void {

    this.getStockType()
  }
  InventorySummaryData: any = [];

  responcestatus: any = ''
  IndividualIS: any = [];
  TotalIS: any = []

  multipleorsingled(block: any, e: any) {


    if (block == 'NU') {


      const index = this.dealType.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.dealType.splice(index, 1);
        if (this.dealType.length == 0) {
          // this.toast.warning('Please Select Atleast One type')
        } else {
          this.getServiceData();
        }
      } else {
        this.dealType = []
        this.dealType.push(e);
        const data = {
          title: 'Inventory Summary',
          stores: this.storeIds,
          toporbottom: this.TotalReport,
          AgeFrom: this.AgeFrom,
          AgeTo: this.AgeTo,
          Month: this.date,
          groups: this.groups,
          dealType: this.dealType.toString(),
          stocktype: this.stocktype,
          statustype: this.statustype, ZeroBalance: this.ZeroBalance.toString(), Wholesale: this.Wholesale,


        };

        this.shared.api.SetHeaderData({ obj: data });
        this.getServiceData();

      }
    }
  }
  getServiceData() {
    this.responcestatus = '';
    this.shared.spinner.show();
    this.GetInventorySummaryReport();
    this.GetInventorySummaryTotalReport();
    localStorage.removeItem('childCalling')
    localStorage.removeItem('childDetailData')
  }
  keys: any = []
  GetInventorySummaryReport() {
    this.IndividualIS = []
    this.shared.spinner.show();
    const obj = {


      "Stores": this.storeIds.toString(),
      "dealtype": this.dealType.toString(),
      "stocktype": this.stocktype == 'All' ? '' : this.stocktype.toString(),
      // "stocktype": '',
      "type": "D",
      "StatusType": this.statustype.length == 8 ? '' : this.statustype.toString(),
      "Wholesale": this.Wholesale.toString().indexOf(',') > 0 ? '' : this.Wholesale.toString(),
      "ZeroBalance": this.ZeroBalance.toString()
    };
    const curl = this.shared.getEnviUrl() + this.shared.common.routeEndpoint + 'GetInventoryEnterpriseSummary';
    this.shared.api.postmethod(this.shared.common.routeEndpoint + "GetInventoryEnterpriseSummary", obj).subscribe(

      (res: { message: any; status: number; response: string | any[] | undefined; }) => {
        const currentTitle = document.title;
        this.shared.api.logSaving(curl, {}, '', res.message, currentTitle);

        if (res.status == 200) {
          if (res.response != undefined) {
            if (res.response.length > 0) {
              this.IndividualIS = res.response;
              if (this.IndividualIS != undefined) {
                if (this.IndividualIS.length > 0) {
                  //console.log(this.IndividualIS);
                  this.keys = Object.keys(this.IndividualIS[0]).slice(7);
                  //console.log(this.keys);
                  this.keys.forEach((val: any, i: any) => {
                    this.IndividualIS.some(function (x: any) {
                      // //console.log(x[val]);
                      if (x[val] != undefined) {
                        x[val] = JSON.parse(x[val])
                      }
                    })

                  });
                  //console.log(this.IndividualIS);

                  let position = this.scrollCurrentposition + 10
                  setTimeout(() => {
                    this.scrollcent.nativeElement.scrollTop = position
                    // //console.log(position);

                  }, 500);
                }


              }
              this.responcestatus = this.responcestatus + 'I';
              this.NoData = false;
              this.combineIndividualandTotal();
            } else {
              // this.toast.error('Empty Response','');
              this.shared.spinner.hide();
              this.NoData = true;
            }
          } else {
            // this.toast.error('Empty Response','');
            this.shared.spinner.hide();
            this.NoData = true;
          }
        } else {
          // this.toast.error(res.status, '');
          this.shared.spinner.hide();
          this.NoData = true;
        }
      },
      (error: any) => {
        // this.toast.error('502 Bad Gate Way Error', '');
        this.shared.spinner.hide();
        this.NoData = true;
      }
    );
  }

  GetInventorySummaryTotalReport() {
    this.TotalIS = []
    this.shared.spinner.show();
    const obj = {


      "Stores": this.storeIds.toString(),
      "dealtype": this.dealType.toString(),
      "stocktype": this.stocktype == 'All' ? '' : this.stocktype.toString(),
      // "stocktype": '',
      "type": "T",
      "StatusType": this.statustype.length == 8 ? '' : this.statustype.toString(),
      "Wholesale": this.Wholesale.toString().indexOf(',') > 0 ? '' : this.Wholesale.toString(),
      "ZeroBalance": this.ZeroBalance.toString()
    };
    this.shared.api.postmethod(this.shared.common.routeEndpoint + "GetInventoryEnterpriseSummary", obj).subscribe(
      (totalres: { status: number; response: any[] | undefined; }) => {
        if (totalres.status == 200) {
          if (totalres.response != undefined) {
            if (totalres.response.length > 0) {
              this.TotalIS = totalres.response.map((v: any) => ({
                store: 'REPORT TOTALS',
                Dealer: '+',
                Data2: [],
                ...v,

              }));
              if (this.TotalIS != undefined) {
                if (this.TotalIS.length > 0) {
                  let keys = Object.keys(this.TotalIS[0]).slice(8);
                  this.keys = Object.keys(this.TotalIS[0]).slice(8);

                  //console.log(keys);

                  //console.log(this.TotalIS);
                  this.keys.forEach((val: any, i: any) => {
                    this.TotalIS.some(function (x: any) {

                      if (x[val] != undefined) {
                        x[val] = JSON.parse(x[val])
                      }

                    });
                  })
                  let position = this.scrollCurrentposition + 10
                  setTimeout(() => {
                    this.scrollcent.nativeElement.scrollTop = position
                    // //console.log(position);

                  }, 500);
                }


              }
              this.responcestatus = this.responcestatus + 'T';
              this.combineIndividualandTotal();
            } else {
              // this.toast.error('Empty Response','');
              this.shared.spinner.hide();
              this.NoData = true;
            }
          } else {
            // this.toast.error('Empty Response','');
            this.shared.spinner.hide();
            this.NoData = true;
          }
        } else {
          // this.toast.error(totalres.status, '');
          this.shared.spinner.hide();
          this.NoData = true;
        }
      },
      (error: any) => {
        // this.toast.error('502 Bad Gate Way Error', '');
        this.shared.spinner.hide();
        this.NoData = true;
      }
    );
  }

  combineIndividualandTotal() {
    if (this.responcestatus == 'IT' || this.responcestatus == 'TI') {

      this.IndividualIS.unshift(this.TotalIS[0]);
      this.InventorySummaryData = this.IndividualIS;
      this.parseData()

      this.shared.spinner.hide();
    } else if (this.responcestatus == 'T') {
      this.InventorySummaryData = this.TotalIS;
      this.parseData()
    } else if (this.responcestatus == 'I') {
      this.InventorySummaryData = this.TotalIS;
      this.parseData()
    } else {
      this.NoData = true;
    }
    console.log(this.InventorySummaryData);
    localStorage.removeItem('DetailsObject');


  }

  nestedsort(property: any, index: any) {
    console.log(property, index);

    this.isDesc = !this.isDesc; //change the direction
    this.column = property;
    let direction = this.isDesc ? 1 : -1;

    this.InventorySummaryData.sort(function (a: any, b: any) {
      // console.log(a[property][0], a[property][0][index]);      
      if (a[property][0][index] < b[property][0][index]) {
        return -1 * direction;
      } else if (a[property][0][index] > b[property][0][index]) {
        return 1 * direction;
      } else {
        return 0;
      }
    });

  }

  parseData() {


  }

  CompleteComponentState: boolean = true;
  subdataindex: any = 0;

  openDetails(id: any, storename: any, range: any, index: any, units: any, block?: any) {
    this.MainGrid = 'N'
    this.storename = storename;
    this.range = range;
    this.totalcount = units;
    this.inventoryDetails = []
    this.getCategoryDealLevel(id, index, range, block)
  }
  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    }
    return false;
  }

  unsubscribing() {
    if (this.reportOpen != undefined) {
      this.reportOpen.unsubscribe()
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

  sorted(property: any) {
    this.isDesc = !this.isDesc; //change the direction
    this.column = property;
    let direction = this.isDesc ? 1 : -1;

    this.inventoryDetails.sort(function (a: any, b: any) {
      if (a[property] < b[property]) {
        return -1 * direction;
      } else if (a[property] > b[property]) {
        return 1 * direction;
      } else {
        return 0;
      }
    });
  }
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
  ngAfterViewInit(): void {
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.shared.common.pageName == 'Inventory Summary') {
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
        if (res.obj.title == 'Inventory Summary') {
          if (res.obj.state == true) {
            if (this.MainGrid == 'Y') {
              this.exportToExcel();
            }
            else {
              this.exportAsXLSX()
            }
          }
        }
      }
    });


    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res: { obj: { title: string; statePrint: boolean; }; }) => {
      if (this.print != undefined) {

        if (res.obj.title == 'Inventory Summary') {
          if (res.obj.statePrint == true) {
            // this.GetPrintData();
          }
        }
      }
    });

    this.Pdf = this.shared.api.getExportToPDFAllReports().subscribe((res: { obj: { title: string; statePDF: boolean; }; }) => {
      if (this.Pdf != undefined) {

        if (res.obj.title == 'Inventory Summary') {
          if (res.obj.statePDF == true) {
            // this.generatePDF();
          }
        }
      }
    });
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res: { obj: { title: string; stateEmailPdf: boolean; }; }) => {
      if (this.email != undefined) {

        if (res.obj.title == 'Inventory Summary') {
          if (res.obj.stateEmailPdf == true) {
            // this.sendEmailData(res.obj.Email, res.obj.notes, res.obj.from);

          }
        }
      }
    });
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
  ngOnDestroy() {
    this.unsubscribing()
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
  // reportOpening(temp: any) {
  //   this.ngmodelactive = this.ngbmodal.open(temp, {
  //     size: 'xl',
  //     backdrop: 'static',
  //   });
  // }
  index = '';
  commentobj = {};




  //-----------Reports---------//
  getStockType() {
    const obj = {}
    this.shared.api.postmethod(this.shared.common.routeEndpoint + "GetInventoryStockTypes", obj).subscribe((res: any) => {
      this.stocktype = res.response;
      console.log(this.stocktype);

    })
  }

  toporbottom: any = ['T'];
  activePopover: number = -1;
  // stocktype: any = [];
  // statustype: any = ['Stock']
  // ZeroBalance: any = ['Y']
  // Wholesale: any = ['N']
  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent) {
    const clickedInside = (event.target as HTMLElement).closest('.dropdown-toggle, .reportstores-card, .timeframe');
    if (!clickedInside) {
      this.activePopover = -1;
    }
  }
  togglePopover(popoverIndex: number) {
    this.activePopover = this.activePopover === popoverIndex ? -1 : popoverIndex;
  }


  multipleorsingle(block: any, e: any) {
    if (block == 'TB') {
      this.toporbottom = [];
      this.toporbottom.push(e);
    }

    if (block == 'NU') {
      this.dealType = [];
      this.dealType.push(e);

      // const index = this.dealType.findIndex((i: any) => i == e);
      // if (index >= 0) {
      //   this.dealType.splice(index, 1);
      // } else {
      //   this.dealType.push(e);
      // }
    }
    if (block == 'WS') {
      // this.Wholesale = [];
      // this.Wholesale.push(e);
      const index = this.Wholesale.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.Wholesale.splice(index, 1);
      } else {
        this.Wholesale.push(e);
      }
    }
    if (block == 'ZB') {
      this.ZeroBalance = [];
      this.ZeroBalance.push(e);
    }
    if (block == 'ST') {
      this.stocktype = [];
      this.stocktype.push(e);
    }
    if (block == 'STT') {

      const index = this.statustype.findIndex((i: any) => i == e);
      if (index >= 0 || this.statustype.length == 8) {
        if (e == 'All') {
          this.statustype = []
        } else {
          this.statustype.splice(index, 1);
        }
      } else {
        if (e == 'All') {
          this.statustype = []
          this.statustype = ['Transit', 'Hold', 'Production', 'Retired', 'Ordered', 'Stock', 'Invoiced', 'Gone']
        } else {
          this.statustype.push(e);
        }
      }
    }
  }
  viewreport() {
    this.activePopover = -1


    if (this.statustype.length == 0) {
    
      this.toast.show('Please Select Atleast One Status Type', 'warning', 'Warning');
    }

    // else if (this.stocktype.length == 0) {
    //   this.toast.warning('Please Select Atleast One Stock Type')
    // }
    // else if(this.AgeTo == 0){
    //   this.toast.warning('Minimum range must be greater then 0')
    // }
    else {
      const data = {
        Reference: 'Inventory Summary',
        // month: this.month,
        // TotalReport: this.toporbottom[0],
        // storeValues: this.selectedstorevalues.toString(),
        AgeFrom: this.AgeFrom,
        AgeTo: this.AgeTo,
        // dataGroupingvalues: this.selectedDataGrouping,
        // groups: this.selectedGroups,
        dealType: this.dealType,
        stocktype: this.stocktype,
        statustype: this.statustype,
        Wholesale: this.Wholesale,
        ZeroBalance: this.ZeroBalance

      };
      //   this.service.SetReports({
      //     obj: data,
      //   });
      //   this.close();
      this.getServiceData()
    }
  }
  details: any = []
  inventoryDetails: any = []
  NoDatacategories: any = ''
  getCategoryDealLevel(id: any, index: any, range: any, block?: any) {
    this.shared.spinner.show()
    const obj = {
      "startdate": "",
      "enddate": "",
      "store": id == undefined ? '' : id,
      "dealtype": this.dealType.toString(),
      "stocktype": '',
      "StatusType": this.statustype,
      "Wholesale": this.Wholesale.toString().indexOf(',') > 0 ? '' : this.Wholesale.toString(),
      "ZeroBalance": this.ZeroBalance.toString(),
      "fromage": block != 'TU' ? (this.keys.length - 1 != index ? range.substring(0, range.indexOf('-')) : range.substring(0, range.indexOf('+'))) : 0,
      "toage": block != 'TU' ? (this.keys.length - 1 != index ? range.substring(range.indexOf('-') + 1, range.indexOf('D')) : '') : 0,
      "PageNumber": 0,
      "PageSize": "1000"

    }
    this.shared.api.postmethod(this.shared.common.routeEndpoint + 'GetEnterpriseSummaryDetails', obj).subscribe(
      res => {
        if (res.status == 200) {
          // this.inventoryDetails = res.response
          this.shared.spinner.hide()

          this.details = res.response.map((v: any) => ({
            ...v, notesView: '+'
          }));
          this.details.some(function (x: any) {
            if (x.Notes != undefined && x.Notes != null) {
              x.Notes = JSON.parse(x.Notes);
            }
          });
          this.inventoryDetails = [
            ...this.inventoryDetails,
            ...this.details,
          ];

          if (this.inventoryDetails.length > 0) {
            this.NoDatacategories = '';
          } else {
            this.NoDatacategories = 'No Data Found!!';
          }
        }
        else {
        
          this.toast.show('Invalid Details.', 'danger', 'Error');
        }
      },
      (error) => {
        //console.log(error);
      })
  }

  notesViewState: boolean = true
  notesView() {
    this.notesViewState = !this.notesViewState
  }
  notesData: any = {}
  // popup: any
  addNotes(data: any, ref: any) {
    //   this.scrollpositionstoring = this.scrollCurrentposition

    //   this.notesData = {
    //     store: this.Salesdetails[0].store,
    //     mainkey: data.stock,
    //     module: 'IN',
    //     apiRoute: 'AddNotesAction'

    //   }
    //   this.popup = this.ngbmodal.open(ref, { size: 'xxl', backdrop: 'static' });
  }
  callLoadingState = 'FL'
  closeNotes(e: any) {
    // this.ngbmodalActive.dismiss()
    this.popup.close()
    if (e == 'S') {
      this.details = [];
      this.inventoryDetails = [];
      this.callLoadingState = 'ANS'
      // this.getCategoryDealLevel()
    }
  }
  toggleView(data: any) {
    if (data.notesView == '+') {
      data.notesView = '-'
    } else {
      data.notesView = '+'
    }
  }
  close() {
    // this.ngbmodel.dismissAll();
  }
  backtoWR() {
    //console.log(this.Salesdetails);
    // this.route.navigateByUrl('InventorySummary')
    this.MainGrid = 'Y'
  }


  getTotal(frontgross: any, colname: any) {
    let total: any = 0
    this.inventoryDetails.some(function (x: any) {
      total += parseFloat(x[colname] == null || x[colname] == undefined ? 0 : x[colname])
    })
    return total
  }


  // getColumnLetter(index: number): string {
  //   return String.fromCharCode(65 + index); // 65 = 'A'
  // }


  getColumnLetter(index: number): string {
    let letters = '';
    while (index > 0) {
      const remainder = (index - 1) % 26;
      letters = String.fromCharCode(65 + remainder) + letters;
      index = Math.floor((index - 1) / 26);
    }
    return letters;
  }
  exportToExcel(): void {
    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('Inventory Summary');

    const titleRow = worksheet.addRow(['Inventory Summary Report']);
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
      { name: 'New Used:', values: this.dealType.toString() },
      { name: 'Wholesale:', values: this.Wholesale.map((item: any) => item === 'Y' ? 'Yes' : 'No').toString() },
      { name: 'W/Balance:', values: this.ZeroBalance.toString() == 'Y' ? 'Yes' : 'No' },

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
    const mainHeaders = ['Store', ...this.keys, 'Totals'];
    const subHeaders = ['Units', 'Value', '% Of $'];
    const topRowIndex = startIndex;
    const subRowIndex = startIndex++;
    const subHeaderRow: string[] = [];
    mainHeaders.forEach((header, index) => {
      if (index === 0) {
        subHeaderRow.push(''); // Store column has no sub-header
      } else if (header === 'Totals') {
        subHeaderRow.push('Units', 'Total', 'Avg Value', 'Units/Day (90 Day Avg)', 'Days Supply');
      } else {
        subHeaderRow.push(...subHeaders);
      }
    });
    worksheet.addRow(subHeaderRow);
    // Merge cells for main headers
    let colIndex = 1;
    mainHeaders.forEach((header, index) => {
      console.log(topRowIndex, subRowIndex, '...................');
      if (index === 0) {
        worksheet.getCell(`${this.getColumnLetter(colIndex)}${topRowIndex}:${this.getColumnLetter(colIndex)}${subRowIndex}`);
        worksheet.getCell(`${this.getColumnLetter(colIndex)}${topRowIndex}`).value = header;
        colIndex++;
      } else if (header === 'Totals') {
        const startCol = this.getColumnLetter(colIndex);
        const endCol = this.getColumnLetter(colIndex + 4);
        worksheet.mergeCells(`${startCol}${topRowIndex}:${endCol}${topRowIndex}`);
        worksheet.getCell(`${startCol}${topRowIndex}`).value = header;
        colIndex += 5;
      } else {
        const startCol = this.getColumnLetter(colIndex);
        const endCol = this.getColumnLetter(colIndex + 2);
        worksheet.mergeCells(`${startCol}${topRowIndex}:${endCol}${topRowIndex}`);
        worksheet.getCell(`${startCol}${topRowIndex}`).value = header;
        colIndex += 3;
      }
    });
    worksheet.getRow(topRowIndex).eachCell(cell => {
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2F5597' } };
    });
    worksheet.getRow(subRowIndex).eachCell(cell => {
      cell.font = { bold: true };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2F5597' } };
    });

    // Dynamic headers
    let dynamicData = ['Units', 'Value', 'percentageof$'];
    let completeData: string[] = [];

    // this.keys.forEach(() => {
    //   completeData = [...completeData, ...dynamicData];
    // });

    this.keys.forEach((key: string) => {
      dynamicData.forEach((metric) => {
        completeData.push(`${key}_${metric}`);
      });
    });


    const bindingHeaders = ['store', ...completeData, 'Units', 'Total', 'Average', 'DayAvg90Days', 'DaysSupply'];
    const currencyFields = ['Value', 'Total', 'Average'];
    console.log(bindingHeaders);

    // Helper function for formatting
    const formatValue = (val: any, key: string): string => {
      if (val == null || val === '') return '-';
      if (currencyFields.includes(key)) return this.cp.transform(val, 'USD', 'symbol', '1.0-0')!;
      if (key === 'percentageof$') return this.cp.transform(val, 'USD', '', '1.2-2')! + '%';
      if (key === 'DayAvg90Days' || key == 'DaysSupply') return this.cp.transform(val, 'USD', '', '1.2-2')!;
      return val;
    };

    // Generate rows dynamically
    for (const info of this.InventorySummaryData) {
      const rowData: any[] = [];

      bindingHeaders.forEach((headerKey) => {
        if (headerKey === 'store') {
          rowData.push(formatValue(info.store, headerKey));
        } else if (headerKey.includes('_')) {
          const [blockKey, metric] = headerKey.split('_');
          const block = info[blockKey]?.[0] || {};
          rowData.push(formatValue(block[metric], metric));
        } else {
          rowData.push(formatValue(info[headerKey], headerKey));
        }
      });
      const dealerRow = worksheet.addRow(rowData);
      bindingHeaders.forEach((key, index) => {
        const cell = dealerRow.getCell(index + 1);
        if (currencyFields.some(field => key.includes(field)) && typeof cell.value === 'string' && cell.value.startsWith('$')) {
          cell.alignment = { horizontal: 'right' };
        } else {
          cell.alignment = { horizontal: 'center' };
        }
      });
    }
    worksheet.columns.forEach((col: any, i: number) => {
      col.width = i === 0 ? 30 : 15;
    });


    workbook.xlsx.writeBuffer().then((buffer: any) => {
      const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      this.shared.exportToExcel(workbook, 'Inventory Summary');
    });
  }



  exportAsXLSX() {

    let localarray = this.inventoryDetails.map((_arrayElement: any) =>
      Object.assign({}, _arrayElement)
    );

    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('Inventory Summary Details');
    let storeValue = 'All Stores';
    // if (
    //   this.storeIds &&
    //   this.storeIds.length > 0 &&
    //   this.storeIds.length !== this.stores.length
    // ) {
      storeValue = this.stores
        .filter((s: any) => this.storeIds.includes(s.ID))
        .map((s: any) => s.storename)
        .join(', ');
    // }
    const filters = [
      { name: 'Store:', values:this.storename != 'REPORT TOTALS' ? this.storename : storeValue },
      { name: 'New Used:', values: this.dealType.toString() },
      { name: 'Wholesale:', values: this.Wholesale.map((item: any) => item === 'Y' ? 'Yes' : 'No').toString() },
      { name: 'W/Balance:', values: this.ZeroBalance.toString() == 'Y' ? 'Yes' : 'No' },
      { name: 'Model:', values: this.range },
    ];
    const ReportFilter = worksheet.addRow(['Inventory Summary Details :']);
    ReportFilter.font = { name: 'Arial', family: 4, size: 10, bold: true };
    let startIndex = 2
    filters.forEach((val: any) => {
      startIndex++
      worksheet.addRow('');
      worksheet.getCell(`A${startIndex}`);
      worksheet.mergeCells(`B${startIndex}:G${startIndex}`);
      worksheet.getCell(`A${startIndex}`).value = val.name;
      worksheet.getCell(`B${startIndex}`).value = val.values
    })
    startIndex++

    worksheet.addRow('');
    startIndex++
    const secondHeader = [
      'Store Name', 'Stock #', 'Age', 'Year', 'Make',
      'Model', 'VIN', 'Color', 'GL Balance', 'Invoice'
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

    const bindingHeaders = ['Dealername', 'stock', 'Age', 'Year', 'Make', 'Model', 'vin', 'Color', 'glbalance', 'Price'];
    const currencyFields: any = ['glbalance', 'Price'];
    let notesCount = 11
    for (const info of localarray) {
      const rowData = bindingHeaders.map(key => {
        const val = info[key];
        return key == 'Dealer' ? ('WesternAuto') : (val === 0 || val == null || val == '' ? '-' : val);
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
        }
      });
      if (info.NotesStatus == 'Y' && this.notesViewState == true) {
        worksheet.mergeCells(notesCount, 1, notesCount, 10);
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
      let maxLength = 15;
      column.width = maxLength + 2;
    });
    workbook.xlsx.writeBuffer().then((buffer: any) => {
      const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      this.shared.exportToExcel(workbook, 'Inventory Summary Details')
    });
  }


}
