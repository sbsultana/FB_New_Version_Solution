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
import { NgbModalModule, NgbModalRef } from '@ng-bootstrap/ng-bootstrap';

@Component({
  selector: 'app-dashboard',
  imports: [SharedModule, BsDatepickerModule, DateRangePicker],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
})
export class Dashboard {

  FromDate: any = '';
  ToDate: any = '';
  minDate!: Date;
  maxDate!: Date;
  DateType: any = 'MTD';
  displaytime: any = '';


  Dates: any = {
    'FromDate': this.FromDate, 'ToDate': this.ToDate, "MaxDate": this.maxDate, 'MinDate': this.minDate, 'DateType': this.DateType, 'DisplayTime': this.displaytime,
    Types: [
      { 'code': 'MTD', 'name': 'MTD' },
      { 'code': 'QTD', 'name': 'QTD' },
      { 'code': 'YTD', 'name': 'YTD' },
      { 'code': 'PYTD', 'name': 'PYTD' },
      { 'code': 'LY', 'name': 'Last Year' },
      { 'code': 'LM', 'name': 'Last Month' },
      { 'code': 'PM', 'name': 'Same Month PY' },
    ]
  }


  IncomeBlock: any[] = []
  ExpensesBlock: any[] = []
  ServiceExpense: any[] = []
  FixedExpense: any[] = []
  Inventory: any[] = []

  IncomeNoData: any = ''
  ExpensesBlockNoData: any = ''
  ServiceNoData: any = ''
  FixedNoData: any = ''
  InventoryNoData: any = ''

  Sequence: any = [
    { 'name': 'Fixed Expenses', seq: 3 },
    { 'name': 'Selling Expenses', seq: 2 },
    { 'name': 'Income', seq: 1 },
  ]

  TotalReport: string = 'T';
  NoData: boolean = false;
  Paytype: any = ['Customerpay_0', 'Warranty_1', 'Internal_2'];
  Department: any = ['Service'];

  responcestatus: any = '';
  groups: any = [7];

  Pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;
  otherstoreid: any = '';
  selectedotherstoreids: any = '';
  selectedRecon: any = [];

  storeIds: any = []
  stores: any = []
  groupName: any = ''
  groupsArray: any = [];
  selectedGroups: any = [];
  storename: any = '';
  Recgroup: any = [];
  Recstores: any = []

  SellingandFixed: any = []
  SellingandfixedData: any = []
  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent) {
    const clickedInside = (event.target as HTMLElement).closest('.dropdown-toggle, .reportstores-card, .timeframe');
    if (!clickedInside) {
      this.activePopover = -1;
    }
  }
  constructor(public shared: Sharedservice, public setdates: Setdates, private comm: common, private cp: CurrencyPipe, private toast: ToastService,) {

    this.shared.setTitle(this.comm.titleName + '-Recon GL Report');

    // if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
    //   this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids.split(',')
    // }
    if (this.shared.common.AccountingReconStores.length > 0) {
      this.groupsArray = this.comm.AccountingReconStores;
      this.stores = this.comm.AccountingReconStores;
      if (this.storeIds == undefined || this.storeIds.length == 0) {
        this.getStoresandGroupsValues('FL')
      }
    }
    this.initializeDates('MTD')
    this.setHeaderData();

  }

  asoftime: any = [];
  ngOnInit() { }

  initializeDates(type: any) {
    let dates: any = this.setdates.setDates(type)
    this.FromDate = dates[0];
    this.ToDate = dates[1];
    localStorage.setItem('time', type);
    this.DateType = type
    this.setDates(this.DateType)

  }
  setHeaderData() {
    // alert('HI')
    const data = {
      title: 'Recon GL Report',
      stores: this.storeIds,
      fromdate: this.FromDate,
      todate: this.ToDate,
      groups: this.groups,
      selectedRecon: this.selectedRecon,

    };
    this.shared.api.SetHeaderData({
      obj: data,
    });
    // this.getServiceData()

  }
  getServiceData() {
    // console.log(this.store, '.........');

    // if (this.store != 0) {
    // console.log(this.store, '.........');

    this.IncomeNoData = ''
    this.ServiceNoData = ''
    this.FixedNoData = ''
    this.InventoryNoData = ''
    this.ExpensesBlockNoData = ''
    this.IncomeBlock = []
    this.ExpensesBlock = []
    this.ServiceExpense = []
    this.FixedExpense = []
    this.Inventory = []

    this.responcestatus = '';

    // this.GetData('')

    this.GetData('')

    // } else {
    //   // this.NoData = true;
    // }
  }
  incomeblockgross = 0;
  GetData(block: any) {
    const obj = {
      startdate: this.FromDate.replaceAll('/', '-'),
      enddate: this.ToDate.replaceAll('/', '-'),
      StoreID: this.storeIds.toString(),
      Dept: '',

    };
    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetServicePartsSummaryGLRecon', obj).subscribe(
      (res) => {
        if (res && res.response && res.response.length > 0) {
          if (block == '') {
            let Expense = res.response
            Expense.some(function (x: any) {
              if (x.Accounts != undefined) {
                x.Accounts = JSON.parse(x.Accounts);
              }
            });
            let data = Expense.reduce(
              (r: any, { category }: any) => {
                if (!r.some((o: any) => o.category == category)) {
                  r.push({
                    category,
                    seq: category == 'Fixed Expenses' ? 3 : (category == 'Selling Expenses' ? 2 : 1),
                    subdata: Expense.filter(
                      (v: any) => v.category == category
                    ),
                    sellingandfixed: Expense.map(function (a: any) {
                      return a.Accounts;
                    })
                  });
                }
                return r;
              },
              []
            ).sort((a: any, b: any) => a.seq - b.seq);


            data.forEach((e: any) => {
              e.subdata.forEach((val: any) => {
                // console.log(val, '...........');

                val.completemerge = val.Accounts.reduce(
                  (r: any, { dept }: any) => {
                    if (!r.some((o: any) => o.dept == dept)) {
                      r.push({
                        dept,
                        DetailsData: val.Accounts.filter(
                          (v: any) => v.dept == dept
                        ),
                      });
                    }
                    return r;
                  },
                  []
                );
                // let income = val.Accounts.reduce((acc: any, curr: any) => {
                //   return acc + (curr.Income || 0)
                // }, 0)
                // let gross = val.Accounts.reduce((acc: any, curr: any) => {
                //   return acc + (curr.gross || 0)
                // }, 0)

                let income = 0;
                let gross = 0;
                let GPPrecent = 0;
                console.log(val, 'Accounts');

                val.Accounts.forEach((x: any) => {

                  income += Number.isNaN(Number(x['Income'])) ? 0 : Number(x['Income']);
                  gross += Number.isNaN(Number(x['gross'])) ? 0 : Number(x['gross']);
                  val.category == 'INCOME' ? (this.incomeblockgross += Number.isNaN(Number(x['gross'])) ? 0 : Number(x['gross'])) : ''
                });
                console.log(gross, income, this.incomeblockgross, 'Gross and Income');

                const result = val.category == 'INCOME' ? (gross / income) * 100 : (gross / this.incomeblockgross) * 100;
                if (!Number.isFinite(result) || Number.isNaN(result)) {
                  GPPrecent = 0;
                } else {
                  GPPrecent = result;
                }
                // console.log(income, gross, this.incomeblockgross, 'Income and Gross');

                val.completemerge.push({
                  dept: 'Total', Income: val.Accounts.reduce((acc: any, curr: any) => {
                    return acc + (curr.Income || 0)
                  }, 0), GPPercent: GPPrecent,
                  gross: val.Accounts.reduce((acc: any, curr: any) => {
                    return acc + (curr.gross || 0)
                  }, 0), DetailsData: [],
                  incomeblockgross: this.incomeblockgross

                })
              })
            })



            this.ExpensesBlock = data;
            // this.SellingandFixed = data.map(function (a: any) {
            //   return a.category;
            // });
            // console.log(this.SellingandFixed);
            // this.SellingandFixed.indexOf('Selling Expenses') >= 0 || this.SellingandFixed.indexOf('Fixed Expenses') >= 0 ?
            this.SellingandfixedData = this.ExpensesBlock[0].sellingandfixed.reduce((acc: any, curr: any) => acc.concat(curr), [])
            // : this.SellingandfixedData = []
            console.log(this.ExpensesBlock, 'Expenses Block');


          }
        }
        else if (res.status == 200) {
          this.emptyData(block)
        } else {
          this.toast.show(res.status, 'danger', 'show');
          this.shared.spinner.hide();
          this.emptyData(block)
        }
      },
      (error) => {
        this.toast.show('502 Bad Gate Way Error', 'danger', 'Error');
        this.shared.spinner.hide();
        this.emptyData(block)
      }
    );
  }

  emptyData(block: any) {
    this.ExpensesBlock ? (this.ExpensesBlock.length > 0 ? this.ExpensesBlockNoData = '' : this.ExpensesBlockNoData = 'No Data Found!!') : this.ExpensesBlockNoData = 'No Data Found!!'

  }


  getTotal(columnname: any, data: any, block?: any) {
    // console.log(data);
    // console.log(columnname, data);
    if (columnname != undefined) {
      if (columnname === 'GPPercent') {
        let income = 0;
        let gross = 0;
        data.forEach((x: any) => {
          income += Number.isNaN(Number(x['Income'])) ? 0 : Number(x['Income']);
          gross += Number.isNaN(Number(x['gross'])) ? 0 : Number(x['gross']);
        });
        // console.log(gross, income, this.incomeblockgross, 'Get Total Function');

        const result = block == 'INCOME' ? (gross / income) * 100 : (gross / this.incomeblockgross) * 100;
        if (!Number.isFinite(result) || Number.isNaN(result)) {
          return 0;
        } else {
          return result;
        }
      }
      else {
        let total = 0
        data.some(function (x: any) {
          total += Number.isNaN(x[columnname]) ? 0 : x[columnname]
        })
        if (Number.isNaN(total)) {
          return 0
        } else {
          return total
        }
      }


    }
    return 0

  }
  getNetProfit(columnname: any) {
    // console.log(data);
    // console.log(columnname, this.SellingandfixedData);
    if (columnname != undefined) {
      if (columnname === 'GPPercent') {
        let income = 0;
        let gross = 0;
        this.SellingandfixedData.forEach((x: any) => {
          income += Number.isNaN(Number(x['Income'])) ? 0 : Number(x['Income']);
          gross += Number.isNaN(Number(x['gross'])) ? 0 : Number(x['gross']);
        });
        console.log(gross, income, this.incomeblockgross, 'Get Total Function');

        const result = (gross / this.incomeblockgross) * 100;
        if (!Number.isFinite(result) || Number.isNaN(result)) {
          return 0;
        } else {
          return result;
        }
      }
      else {
        let total = 0
        this.SellingandfixedData.some(function (x: any) {
          total += Number.isNaN(x[columnname]) ? 0 : x[columnname]
        })
        if (Number.isNaN(total)) {
          return 0
        } else {
          return total
        }
      }


    }
    return 0

  }



  popupReference!: NgbModalRef;
  popupvalues: any = { FIN: '', Acctno: '', Dept: '', Store: '', Storename: '', subtype: '', department: '' }
  ReconDetails: any = []
  ReconDetailsNoData: any = ''
  reconaccountpopup(tmp: any, FIN: any, Acctno: any, Dept: any, Store: any, StoreName: any, subtype: any, department: any) {
    this.popupvalues.FIN = FIN;
    this.popupvalues.Acctno = Acctno;
    this.popupvalues.Dept = Dept;
    this.popupvalues.Store = Store;
    this.popupvalues.Storename = StoreName;
    this.popupvalues.subtype = subtype;
    this.popupvalues.department = department;

    this.ReconDetailsNoData = '';

    this.ReconDetails = []
    this.popupReference = this.shared.ngbmodal.open(tmp, { size: 'xl', backdrop: 'static', keyboard: true, centered: true, modalDialogClass: 'custom-modal' })
    console.log(FIN, Acctno, Dept, Store, subtype);

    this.GetReconDetails(FIN, Acctno, Dept, Store, subtype)
  }

  GetReconDetails(FIN: any, Acctno: any, Dept: any, Store: any, subtype: any) {
    const obj = {
      "startdate": this.FromDate,
      "enddate": this.ToDate,
      "StoreID": this.storeIds.toString(),
      "account": Acctno ? Acctno : '',
      "dept": Dept,
      "subtype": subtype,
      "category": FIN
      // "finsummary": FIN
    }
    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetServicePartsSummaryGLReconDetailsV1', obj).subscribe(
      (res) => {
        if (res.response && res.response.length > 0) {
          this.ReconDetails = res.response
          this.ReconDetailsNoData = ''
        } else {
          this.ReconDetailsNoData = 'No Data Found'
        }
      },
      (error) => {
        this.ReconDetailsNoData = 'No Data Found'
      })
  }
  closePopup() {
    if (this.popupReference) {
      this.popupReference.close();
    }
  }

  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    } else if (value < 0) {
      return false;
    }
    return true;
  }


  isDesc: boolean = false;
  column: string = 'CategoryName';
  sort(property: any, data: any) {
    this.isDesc = !this.isDesc; //change the direction
    this.column = property;
    let direction = this.isDesc ? 1 : -1;
    // //console.log(property)
    data.sort(function (a: any, b: any) {
      if (a[property] < b[property]) {
        return -1 * direction;
      } else if (a[property] > b[property]) {
        return 1 * direction;
      } else {
        return 0;
      }
    });
  }

  ngAfterViewInit() {

    this.shared.api.getAccountingAllStores().subscribe((res: any) => {
      if (this.comm.pageName == 'Recon GL Report') {
        if (res.obj.storesData != undefined) {
          this.stores = res.obj.storesData;
          this.storeIds = []
          if (this.storeIds == undefined || this.storeIds.length == 0) {
            this.getGroupBaseStores(this.groups, 'FL')
            console.log(this.stores, 'S');

          }

        }
      }
    })

    this.shared.api.getAllStores().subscribe((res: any) => {
      if (this.comm.pageName == 'Recon GL Report') {
        if (res.obj.storesData != undefined) {
          this.groupsArray = res.obj.storesData;
          this.Recstores = res.obj.storesData.filter((v: any) => v.sg_id == 7)[0].Stores
          console.log(this.stores, this.Recstores, 'Stores Rec Stores');

          // this.groupsBindingData()
          if (this.storeIds == undefined || this.storeIds.length == 0) {
            this.getGroupBaseStores(this.groups, 'FL')
          }

        }
      }
    })

    this.excel = this.shared.api.getExportToExcelAllReports().subscribe((res) => {
      if (this.excel != undefined) {
        if (res.obj.title == 'Recon GL Report') {
          if (res.obj.state == true) {
            this.exportToExcel();
          }
        }
      }
    });
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res) => {
      if (this.email != undefined) {
        if (res.obj.title == 'Recon GL Report') {
          if (res.obj.stateEmailPdf == true) {
            // this.sendEmailData(res.obj.Email, res.obj.notes, res.obj.from);
          }
        }
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res) => {
      if (this.print != undefined) {
        if (res.obj.title == 'Recon GL Report') {
          if (res.obj.statePrint == true) {
            // this.GetPrintData();
          }
        }
      }
    });
    this.Pdf = this.shared.api.getExportToPDFAllReports().subscribe((res) => {
      if (this.Pdf != undefined) {
        if (res.obj.title == 'Recon GL Report') {
          if (res.obj.statePDF == true) {
            // this.generatePDF();
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

  activePopover: number = -1;
  togglePopover(popoverIndex: number) {
    if (this.activePopover === popoverIndex) {
      this.activePopover = -1;
    } else {
      this.activePopover = popoverIndex;
    }
  }

  individualgroups(e: any) {
    this.groups = []
    this.groups.push(e.sg_id);
    // this.groupName = this.groupsArray.filter((val: any) => val.sg_id == this.groups[0])[0].sg_name;
    this.groupName = 'Recon Centers';

    this.getGroupBaseStores(this.groups.toString(), 'FG')
  }

  spinnerLoader: boolean = false;
  getGroupBaseStores(id: any, block: any) {
    this.spinnerLoader = true
    // this.stores = this.comm.groupsandstores.filter((v: any) => v.sg_id == id)[0].Stores;

    this.stores = this.comm.AccountingReconStores && this.comm.AccountingReconStores.length > 0 ? this.comm.AccountingReconStores : []
    this.Recstores = this.comm.ReconStores && this.comm.ReconStores.length > 0 ? this.comm.ReconStores.filter((v: any) => v.sg_id == 7)[0].Stores : []

    console.log(this.stores, this.Recstores, 'Stores Rec Stores');

    this.spinnerLoader = false
    this.storeIds = []
    if (block == 'FG') {
      this.storeIds = this.stores.map(function (a: any) {
        return a.companyid;
      });
      this.groupName = 'Recon Centers';
    }
    if (block == 'FL') {
      this.getStoresandGroupsValues('FL')
    }
  }
  getStoresandGroupsValues(type?: any) {
    // alert('Hello')
    let data = JSON.parse(localStorage.getItem('UserDetails')!)
    let dupStores = this.stores.map(function (a: any) {
      return a.companyid;
    });
    let duptokenstores: any = [];
    // type == 'FL' ? JSON.parse(localStorage.getItem('UserDetails')!).Store_Ids.indexOf(',') > 0 ? duptokenstores = JSON.parse(localStorage.getItem('UserDetails')!).Store_Ids.split(',') : duptokenstores.push(JSON.parse(localStorage.getItem('UserDetails')!).Store_Ids) : ''
    type == 'FL' ? JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids.indexOf(',') > 0 ? duptokenstores = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids.split(',') : duptokenstores.push(JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids) : ''
    const intersection = dupStores.filter((element: any) => duptokenstores.includes(element));
    console.log(intersection, duptokenstores, dupStores);

    this.storeIds = []
    type == 'FL' ? this.storeIds = intersection :
      this.storeIds = this.stores.map(function (a: any) {
        return a.companyid;
      });
    console.log(this.storeIds, '.............');


    this.setHeaderData()
    if (this.storeIds && this.storeIds.length > 0 && this.FromDate != '' && this.ToDate != '') {
      this.getServiceData()
    }
    if (intersection.length == 0) {
      this.IncomeNoData = 'Please Select Store'
      this.ServiceNoData = 'Please Select Store'
      this.FixedNoData = 'Please Select Store'
      this.InventoryNoData = 'Please Select Store'
      this.ExpensesBlockNoData = 'Please Select Store'
      this.IncomeBlock = []
      this.ExpensesBlock = []
      this.ServiceExpense = []
      this.FixedExpense = []
      this.Inventory = []
      this.responcestatus = '';
    }
    if (this.stores.length == this.storeIds.length) {
      // this.groupName = this.groupsArray.filter((val: any) => val.sg_id == this.groups[0])[0].sg_name;
      this.groupName = 'Recon Centers';

    }
    if (this.storeIds.length == 1) {
      this.storename = this.stores.filter((val: any) => val.companyid == this.storeIds.toString())[0].storename
    }
  }
  individualStores(e: any) {
    const index = this.storeIds.findIndex((i: any) => i == e.companyid);
    if (index >= 0) {
      this.storeIds.splice(index, 1);
    } else {
      this.storeIds.push(e.companyid);
    }
    if (this.storeIds.length == 1) {
      this.storename = this.stores.filter((val: any) => val.companyid == this.storeIds.toString())[0].storename
    }
  }


  allstores(state: any) {
    if (state == 'N') {
      this.storeIds = [];
    } else if (state == 'Y') {
      this.storeIds = this.stores.map(function (a: any) {
        return a.ID;
      });
      this.groupName = 'Recon Centers';
    }

  }

  ReconindividualStores(e: any) {
    this.selectedRecon = []
    this.selectedRecon.push(e)
    if (this.selectedRecon.length == 2) {
      this.storeIds = this.stores.map(function (a: any) { return a.companyid; })
    }
    else if (this.selectedRecon.length == 0) {
      this.storeIds = []
    } else {
      if (this.selectedRecon[0] == 16) {
        this.storeIds = this.stores.filter((e: any) => e.companyid != '40' && e.companyid != '41' && e.companyid != '42')
          .map(function (a: any) {
            return a.companyid;
          })
        console.log(this.storeIds, '.................');

      } else {
        this.storeIds = ['40', '41', '42']
      }
    }

  }
  viewreport() {
    this.activePopover = -1
    if (this.storeIds.length == 0) {
      this.toast.show('Please select atleast one store', 'warning', 'Warning');
    }
    else {
      // const data = {
      //   Reference: 'Recon GL Report',
      //   storeValues: this.storeIds.toString(),
      //   groups: this.groups.toString(),
      //   selectedrecon: this.selectedRecon
      // };
      // this.shared.api.SetReports({
      //   obj: data,
      // });
      this.setHeaderData();
      this.getServiceData()
    }
  }


  viewDeal(dealData: any) {
    // const modalRef = this.ngbmodal.open(DealRecapComponent, { size: 'md', windowClass: 'compModal' });
    // modalRef.componentInstance.data = { dealno: dealData.dealno, storeid: dealData.dealerid, stock: dealData.Stock, vin: dealData.VIN, custno: dealData?.ad_custid }; // Pass data to the modal component    
    // modalRef.result.then((result) => {
    //   console.log(result); // Handle modal close result
    // }, (reason) => {
    //   console.log(`Dismissed: ${reason}`); // Handle dismiss reason
    // });
  }

  handler(value: string): void {
    if ('onShown' === value) {
    }
    if ('onHidden' === value) {
      (<HTMLInputElement>document.getElementById('DateOfBirth')).click();
    }
  }


  exportToExcel() {
    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('Recon GL Report');
    worksheet.views = [{ showGridLines: false }];

    worksheet.addRow('');
    const titleRow = worksheet.addRow(['Recon GL Report']);
    titleRow.eachCell((cell: any) => {
      cell.alignment = { indent: 1, vertical: 'middle', horizontal: 'left' };
    });
    titleRow.font = { name: 'Arial', family: 4, size: 12, bold: true };
    titleRow.worksheet.mergeCells('A2', 'E2');
    worksheet.addRow('');

    const DateToday = this.shared.datePipe.transform(new Date(), 'MM/dd/yyyy h:mm:ss a');
    const DATE_EXTENSION = this.shared.datePipe.transform(new Date(), 'MMddyyyy');
    worksheet.addRow([DateToday]).font = { name: 'Arial', family: 4, size: 9 };

    worksheet.addRow([]);
    worksheet.addRow(['Report Controls:']).font = { bold: true };



    const timeFrameRow = worksheet.addRow(['Timeframe:', `${this.FromDate} to ${this.ToDate}`]);
    timeFrameRow.font = { name: 'Arial', size: 9 };


    // SECTION 1: IncomeBlock
    worksheet.addRow([]);
    const grandTotalRow1 = worksheet.addRow(['Income', '']);
    grandTotalRow1.font = { name: 'Arial', bold: true, size: 11 };
    grandTotalRow1.getCell(2).numFmt = '$#,##0.00';
    grandTotalRow1.eachCell((cell) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '' },
      };
      cell.alignment = { horizontal: 'right' };

    });

    worksheet.addRow([]);

    const allDetailCells: string[] = [];

    for (const deptItem of this.IncomeBlock) {
      const deptRow = worksheet.addRow([`${deptItem.Dept}`, deptItem.Postingamount || '']);
      deptRow.font = { name: 'Arial', bold: true, size: 10 };
      deptRow.getCell(2).numFmt = '$#,##0.00';
      deptRow.eachCell((cell) => {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: '' },
        };
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' },
        };
      });

      let startRow = (worksheet.lastRow?.number ?? 0) + 1;

      for (const detail of deptItem.DetailsData) {
        const detailRow = worksheet.addRow([detail.SubtypeDetail, detail.Postingamount]);
        detailRow.font = { name: 'Arial', size: 9 };
        detailRow.getCell(2).numFmt = '$#,##0.00';
        detailRow.eachCell((cell) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: '' },
          };
          cell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' },
          };
        });
      }

      let endRow = worksheet.lastRow?.number ?? 0;

      const sellingTotalRow = worksheet.addRow(['Total', { formula: `SUM(B${startRow}:B${endRow})` }]);
      sellingTotalRow.font = { name: 'Arial', bold: true, size: 9 };
      sellingTotalRow.getCell(2).numFmt = '$#,##0.00';
      sellingTotalRow.eachCell((cell) => {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: '' },
        };
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' },
        };
      });

      allDetailCells.push(`B${sellingTotalRow.number}`);
      worksheet.addRow([]);
    }

    // SECTION 2: ServiceExpense
    if (this.ServiceExpense?.length) {
      for (const categoryBlock of this.ServiceExpense) {
        worksheet.addRow([]);
        const grandTotalRow1 = worksheet.addRow(['Selling Expenses', '']);
        grandTotalRow1.font = { name: 'Arial', bold: true, size: 11 };
        grandTotalRow1.getCell(2).numFmt = '$#,##0.00';
        grandTotalRow1.eachCell((cell) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: '' },
          };
          cell.alignment = { horizontal: 'right' };

        });

        for (const storeData of categoryBlock.subdata) {
          const storeRow = worksheet.addRow([storeData.StoreName, '']);
          storeRow.font = { name: 'Arial', bold: true };
          storeRow.getCell(1).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: '' },

          };
          storeRow.eachCell((cell) => {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: '' },
            };
            cell.border = {
              top: { style: 'thin' },
              left: { style: 'thin' },
              bottom: { style: 'thin' },
              right: { style: 'thin' },
            };
          });

          const deptData = storeData.completemerge || [];

          for (const deptItem of deptData) {
            const deptRow = worksheet.addRow([deptItem.Dept, deptItem.postingamount || '']);
            deptRow.font = { name: 'Arial', bold: true, size: 10 };
            deptRow.getCell(2).numFmt = '$#,##0.00';
            deptRow.eachCell((cell) => {
              cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: '' },
              };
              cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' },
              };
            });

            const startRow = (worksheet.lastRow?.number ?? 0) + 1;

            for (const detail of deptItem.DetailsData || []) {
              const detailRow = worksheet.addRow([detail.SubtypeDetail, detail.postingamount]);
              detailRow.font = { name: 'Arial', size: 9 };
              detailRow.getCell(2).numFmt = '$#,##0.00';
              detailRow.eachCell((cell) => {
                cell.fill = {
                  type: 'pattern',
                  pattern: 'solid',
                  fgColor: { argb: '' },
                };
                cell.border = {
                  top: { style: 'thin' },
                  left: { style: 'thin' },
                  bottom: { style: 'thin' },
                  right: { style: 'thin' },
                };
              });
            }

            const endRow = worksheet.lastRow?.number ?? 0;
            if (startRow <= endRow) {
              const totalRow = worksheet.addRow([
                'Total',
                { formula: `SUM(B${startRow}:B${endRow})` },
              ]);
              totalRow.font = { name: 'Arial', bold: true, size: 9 };
              totalRow.getCell(2).numFmt = '$#,##0.00';
              totalRow.eachCell((cell) => {
                cell.fill = {
                  type: 'pattern',
                  pattern: 'solid',
                  fgColor: { argb: '' },
                };
                cell.border = {
                  top: { style: 'thin' },
                  left: { style: 'thin' },
                  bottom: { style: 'thin' },
                  right: { style: 'thin' },
                };
              });
            }

            worksheet.addRow([]);
          }
        }
      }
      // SECTION 3: FixedExpenses
      if (this.FixedExpense?.[0]?.category === 'FixedExpenses' && this.FixedExpense[0].subdata?.length) {
        worksheet.addRow([]);
        const grandTotalRow = worksheet.addRow(['Fixed Expenses', '']);
        grandTotalRow.font = { name: 'Arial', bold: true, size: 11 };
        grandTotalRow.eachCell((cell) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: '' },
          };
          cell.alignment = { horizontal: 'right' };

        });

        for (const storeData of this.FixedExpense[0].subdata) {
          const storeRow = worksheet.addRow([storeData.StoreName, '']);
          storeRow.font = { name: 'Arial', bold: true };
          storeRow.getCell(1).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: '' },
          };
          storeRow.eachCell((cell) => {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: '' },
            };
            cell.border = {
              top: { style: 'thin' },
              left: { style: 'thin' },
              bottom: { style: 'thin' },
              right: { style: 'thin' },
            };
          });
          const dept = storeData.Accounts?.[0]?.Dept || '-';
          const deptRow = worksheet.addRow([dept, '']);
          deptRow.font = { name: 'Arial', bold: true, size: 10 };
          deptRow.eachCell((cell) => {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: '' },
            };
            cell.border = {
              top: { style: 'thin' },
              left: { style: 'thin' },
              bottom: { style: 'thin' },
              right: { style: 'thin' },
            };
          });

          const startRow = (worksheet.lastRow?.number ?? 0) + 1;

          for (const acc of storeData.Accounts || []) {
            const detailRow = worksheet.addRow([acc.SubtypeDetail, acc.postingamount]);
            detailRow.font = { name: 'Arial', size: 9 };
            detailRow.getCell(2).numFmt = '$#,##0.00';
            detailRow.eachCell((cell) => {
              cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: '' },
              };
              cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' },
              };
            });
          }

          const endRow = worksheet.lastRow?.number ?? 0;

          if (startRow <= endRow) {
            const totalRow = worksheet.addRow([
              'Total',
              { formula: `SUM(B${startRow}:B${endRow})` },
            ]);
            totalRow.font = { name: 'Arial', bold: true, size: 9 };
            totalRow.getCell(2).numFmt = '$#,##0.00';
            totalRow.eachCell((cell) => {
              cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: '' },
              };
              cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' },
              };
            });
          }

          worksheet.addRow([]);

        }
      }
      // SECTION 4: Inventory
      //   if (this.Inventory?.[0]?.category === 'Inventory' && this.Inventory[0].subdata?.length) {
      //     worksheet.addRow([]);
      //     const grandTotalRow = worksheet.addRow(['Inventory', '']);
      //     grandTotalRow.font = { name: 'Arial', bold: true, size: 11 };
      //     grandTotalRow.eachCell((cell) => {
      //       cell.fill = {
      //         type: 'pattern',
      //         pattern: 'solid',
      //         fgColor: { argb: '' },
      //       };
      //       cell.alignment = { horizontal: 'right' };
      //     });

      //     for (const storeData of this.Inventory[0].subdata) {
      //       const storeRow = worksheet.addRow([storeData.StoreName,'']);
      //       storeRow.font = { name: 'Arial', bold: true };
      //       storeRow.getCell(1).fill = {
      //         type: 'pattern',
      //         pattern: 'solid',
      //         fgColor: { argb: '' },
      //       };
      //       storeRow.eachCell((cell) => {
      //         cell.fill = {
      //           type: 'pattern',
      //           pattern: 'solid',
      //           fgColor: { argb: '' },
      //         };
      //         cell.border = {
      //           top: { style: 'thin' },
      //           left: { style: 'thin' },
      //           bottom: { style: 'thin' },
      //           right: { style: 'thin' },
      //         };
      //       });
      //       const dept = storeData.Accounts?.[0]?.Dept || '-';
      //       const deptRow = worksheet.addRow([dept, '']);
      //       deptRow.font = { name: 'Arial', bold: true, size: 10 };
      //       deptRow.eachCell((cell) => {
      //         cell.fill = {
      //           type: 'pattern',
      //           pattern: 'solid',
      //           fgColor: { argb: '' },
      //         };
      //         cell.border = {
      //           top: { style: 'thin' },
      //           left: { style: 'thin' },
      //           bottom: { style: 'thin' },
      //           right: { style: 'thin' },
      //         };
      //       });

      //       const startRow = (worksheet.lastRow?.number ?? 0) + 1;

      //       for (const acc of storeData.Accounts || []) {
      //         const detailRow = worksheet.addRow([acc.SubtypeDetail, acc.postingamount]);
      //         detailRow.font = { name: 'Arial', size: 9 };
      //         detailRow.getCell(2).numFmt = '$#,##0.00';
      //         detailRow.eachCell((cell) => {
      //           cell.fill = {
      //             type: 'pattern',
      //             pattern: 'solid',
      //             fgColor: { argb: '' },
      //           };
      //           cell.border = {
      //             top: { style: 'thin' },
      //             left: { style: 'thin' },
      //             bottom: { style: 'thin' },
      //             right: { style: 'thin' },
      //           };
      //         });
      //       }
      //       const endRow = worksheet.lastRow?.number ?? 0;

      //       if (startRow <= endRow) {
      //         const totalRow = worksheet.addRow([
      //           'Total',
      //           { formula: `SUM(B${startRow}:B${endRow})` },
      //         ]);
      //         totalRow.font = { name: 'Arial', bold: true, size: 9 };
      //         totalRow.getCell(2).numFmt = '$#,##0.00';
      //         totalRow.eachCell((cell) => {
      //           cell.fill = {
      //             type: 'pattern',
      //             pattern: 'solid',
      //             fgColor: { argb: '' },
      //           };
      //           cell.border = {
      //             top: { style: 'thin' },
      //             left: { style: 'thin' },
      //             bottom: { style: 'thin' },
      //             right: { style: 'thin' },
      //           };
      //         });
      //       }

      //       worksheet.addRow([]);

      // }
      //   }
    }


    // Adjust column widths
    worksheet.getColumn(1).width = 25;
    worksheet.getColumn(2).width = 20;


    this.shared.exportToExcel(workbook, 'Recon GL Report V1_' + DATE_EXTENSION + '.xlsx');

  }


  Details_ExportAsXLSX() {
    let storeNames: any = [];

    let FSDetailsData = this.ReconDetails.map((_arrayElement: any) =>
      Object.assign({}, _arrayElement)
    );
    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('Recon GL Report');
    worksheet.views = [
      {
        state: 'frozen',
        ySplit: 10, // Number of rows to freeze (2 means the first two rows are frozen)
        topLeftCell: 'A11', // Specify the cell to start freezing from (in this case, the third row)
        showGridLines: false,
      },
    ];
    worksheet.addRow('');
    const titleRow = worksheet.addRow(['Recon GL Report']);
    titleRow.eachCell((cell, number) => {
      cell.alignment = { indent: 1, vertical: 'middle', horizontal: 'left' };
    });
    titleRow.font = { name: 'Arial', family: 4, size: 12, bold: true };
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

    const ReportFilter = worksheet.addRow(['Storename :', this.popupvalues.Storename]);
    ReportFilter.font = { name: 'Arial', family: 4, size: 10, bold: true };

    const Type = worksheet.getCell('A6');
    Type.value = 'Type :';
    Type.font = { name: 'Arial', family: 4, size: 9, bold: true };

    const type = worksheet.getCell('B6');
    type.value = this.popupvalues.FIN;
    type.font = { name: 'Arial', family: 4, size: 9 };
    type.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };

    // Row 7: Account Number
    const DateMonth = worksheet.getCell('A7');
    DateMonth.value = 'Account Number :';
    DateMonth.font = { name: 'Arial', family: 4, size: 9, bold: true };

    const date = worksheet.getCell('B7');
    date.value = this.popupvalues.Acctno;
    date.font = { name: 'Arial', family: 4, size: 9 };
    date.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };

    // worksheet.mergeCells('B7', 'K9');
    // const Store = worksheet.getCell('A8');
    // Store.value = 'Store :'
    // const stores = worksheet.getCell('B8');
    // stores.value = this.ExcelStoreNames == 0
    // ? 'All Stores'
    // : this.ExcelStoreNames == null
    // ? '-'
    // : this.ExcelStoreNames.toString().replaceAll(',', ', ');
    // stores.font = { name: 'Arial', family: 4, size: 9 };
    // store.alignment = { vertical: 'middle', horizontal: 'left',wrapText:true};
    // Store.font = {
    //   name: 'Arial',
    //   family: 4,
    //   size: 9,
    //   bold: true,
    // };

    worksheet.addRow('');

    let Headings = [
      'Date',
      'Refer',
      'Posting Amount',
      'Detail Description'
    ];
    const headerRow = worksheet.addRow(Headings);
    headerRow.height = 20;
    headerRow.font = {
      name: 'Arial',
      family: 4,
      size: 8,
      bold: true,
      color: { argb: 'FFFFFF' },
    };
    headerRow.alignment = { indent: 1, vertical: 'middle', horizontal: 'center' };
    headerRow.eachCell((cell, number) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '2a91f0' },
        bgColor: { argb: 'FF0000FF' },
      };
      cell.border = { right: { style: 'dotted' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });

    for (const d of FSDetailsData) {
      const Data1 = worksheet.addRow([
        d.accountingdate == '' ? '-' : d.accountingdate == null ? '-' : d.accountingdate,
        d.refer == '' ? '-' : d.refer == null ? '-' : d.refer,
        d.postingamount == '' ? '-' : d.postingamount == null ? '-' : '$' + d.postingamount,
        // d.accounttype == '' ? '-' : d.accounttype == null ? '-' : d.accounttype,
        // d.accountdescription == '' ? '-' : d.accountdescription == null ? '-' : d.accountdescription,
        // d.GroupDesc == '' ? '-' : d.GroupDesc == null ? '-' : d.GroupDesc,
        // d.SubGroupDesc == '' ? '-' : d.SubGroupDesc == null ? '-' : d.SubGroupDesc,
        d.detaildescription == '' ? '-' : d.detaildescription == null ? '-' : d.detaildescription,
        // d.postingamount == '' ? '-' : d.postingamount == null ? '-' : '$'+ '      ' +d.postingamount,
      ]);
      Data1.font = { name: 'Arial', family: 4, size: 8 };

      // Set alignment for each cell individually
      Data1.getCell(1).alignment = { indent: 1, vertical: 'middle', horizontal: 'center' };
      Data1.getCell(2).alignment = { indent: 1, vertical: 'middle', horizontal: 'center' };
      Data1.getCell(3).alignment = { indent: 1, vertical: 'middle', horizontal: 'right' };
      Data1.getCell(4).alignment = { indent: 1, vertical: 'middle', horizontal: 'center' };
      Data1.getCell(5).alignment = { indent: 1, vertical: 'middle', horizontal: 'left' };
      Data1.getCell(6).alignment = { indent: 1, vertical: 'middle', horizontal: 'left' };
      Data1.getCell(7).alignment = { indent: 1, vertical: 'middle', horizontal: 'left' };
      Data1.getCell(8).alignment = { indent: 1, vertical: 'middle', horizontal: 'center' };
      Data1.getCell(4).numFmt = '$#,##0';
      Data1.eachCell((cell, number) => {
        cell.border = { right: { style: 'dotted' } };
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
    }
    worksheet.eachRow((row, rowIndex) => {
      row.eachCell((cell, colIndex) => {
        if (rowIndex > 1 && rowIndex < 9) {
          if (colIndex === 1) {
            cell.alignment = {
              horizontal: 'left',
              vertical: 'middle',
              indent: 1,
            };
          }
        }
      });
    });
    worksheet.getColumn(1).width = 25;
    worksheet.getColumn(2).width = 15;
    worksheet.getColumn(3).width = 20;
    worksheet.getColumn(4).width = 15;
    worksheet.getColumn(5).width = 25;
    worksheet.getColumn(6).width = 25;
    worksheet.getColumn(7).width = 30;
    worksheet.getColumn(8).width = 50;
    worksheet.addRow([]);

    this.shared.exportToExcel(workbook, 'Recon GL Report Composite_' + DATE_EXTENSION);

  }

}