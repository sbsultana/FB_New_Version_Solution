



import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { Component, ElementRef, ViewChild } from '@angular/core';
import { Subscription } from 'rxjs';
import { Setdates } from '../../../../Core/Providers/SetDates/setdates';
import { SalesgrossReports } from '../salesgross-reports/salesgross-reports';
import { SalesgrossDetails } from '../salesgross-details/salesgross-details';

@Component({
  selector: 'app-dashboard',
  imports: [SharedModule, SalesgrossReports],
  standalone: true,
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss'
})
export class Dashboard {
  @ViewChild('printSection') printSection!: ElementRef;
  FromDate: any = '';
  ToDate: any = '';
  SalesData: any = [];
  IndividualSalesGross: any = [];
  TotalSalesGross: any = [];
  IndividualSalesGrossBackGross: any = [];
  TotalSalesGrossBackGross: any = [];
  BackGross: any = [];
  DateType: any = 'MTD';
  GridView = 'Global';
  TotalReport: any = 'B';
  TotalSortPosition: any = 'B';
  storeIds: any = '';
  salesPersonId: any = '0';
  salesManagerId: any = '0';
  financeManagerId: any = '0';
  dealType: any = ['New', 'Used'];
  saleType: any = ['Retail', 'Lease', 'Misc', 'Special Order'];
  dealStatus: any = ['Delivered', 'Capped', 'Finalized'];
  target: any = [];
  source: any = [];
  includecharge: any = [];
  
  pack: any = [];
  path1: any = 'store';
  path2: any = 'ad_dealtype';
  path3: any = '';
  path1name: any = 'All Dealerships';
  path2name: any = 'New/Used';
  path3name: any = '';
  path1id: any = 1;
  path2id: any = 2;
  path3id: any = '';
  CurrentDate = new Date();
  NoData: boolean = false;
  responcestatus: string = '';
  groups: any = 0;
  acquisition: any = ['All'];
  otherstoreid: any = '';
  selectedotherstoreids: any = '';
  actionType:any=''
  header: any = [
    {
      type: 'Bar',
      storeIds: this.storeIds,
      fromDate: this.FromDate,
      toDate: this.ToDate,
      ReportTotal: this.TotalReport, groups: this.groups,
      sp: this.salesPersonId,
      sm: this.salesManagerId,
      fm: this.financeManagerId,
      as: this.acquisition,
      gridview: this.GridView, otherstoreids: this.otherstoreid, selectedotherstoreids: this.selectedotherstoreids, datevaluetype: this.DateType, dealType: this.dealType,
      saleType: this.saleType,
      dealStatus: this.dealStatus,
    },
  ];


  reportOpenSub!: Subscription;
  reportGetting!: Subscription;
  Pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;

  constructor(public shared: Sharedservice, public setdates: Setdates) {

    this.shared.setTitle(this.shared.common.titleName + '-Sales Gross')

    if (typeof window !== 'undefined') {
      console.log(localStorage.getItem('stime'));
      if (localStorage.getItem('flag') == 'V') {
        this.storeIds = [];
        console.log(JSON.parse(localStorage.getItem('userInfo')!), JSON.parse(localStorage.getItem('userInfo')!).user_Info, 'Widget Stores............');
        this.groups = JSON.parse(localStorage.getItem('userInfo')!).groupid
        JSON.parse(localStorage.getItem('userInfo')!).store.indexOf(',') > 0 ?
          this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).store.split(',') :
          this.storeIds.push(JSON.parse(localStorage.getItem('userInfo')!).store)   
          localStorage.setItem('flag','M')    
      } else {
        if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
          console.log(JSON.parse(localStorage.getItem('userInfo')!).user_Info);          
          this.groups = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Preferences
          this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids.split(',')
        }
      }

      // this.storeIds = '1,2'

      if (localStorage.getItem('stime') != null) {
        let stime = localStorage.getItem('stime');
        if (stime != null && stime != '') {
          this.setDates(stime)
          this.DateType = stime
        }
      } else {
        this.setDates('MTD')
        this.DateType = 'MTD'
      }
      // localStorage.setItem('stime', 'MTD')
      this.getPeopleList()
    }
  }

  ngOnInit(): void {
    // localStorage.setItem('time', 'MTD');
    // this.reportOpenSub.unsubscribe()
    // this.reportGetting.unsubscribe()
    // this.Pdf.unsubscribe()
    // this.print.unsubscribe()
    // this.email.unsubscribe()
    // this.excel.unsubscribe()
  }

  Favreports: any = [];
  getPeopleList() {
    // const obj = {
    //   "Stores": this.storeIds
    // }
    // this.shared.api.postmethod(this.shared.common.routeEndpoint+'GetPeopleList',obj).subscribe((res:any)=>{
    //     if(res.status == 200){
    //       console.log(res.response);
    //       if(res.response.length >0){
    //         this.salesPersonId=res.response[0].SalesPersons;
    //         this.financeManagerId=res.response[0].FIManager;
    //         this.salesManagerId= res.response[0].SalesManager;
    const data = {
      title: 'Sales Gross',
      path1: this.path1name,
      path2: this.path2name,
      path3: this.path3name,
      path1id: this.path1id,
      path2id: this.path2id,
      path3id: this.path3id,
      stores: this.storeIds,
      salespresons: '',
      salesmanagers: '',
      financemanagers: '',
      dealType: this.dealType,
      saleType: this.saleType,
      dealStatus: this.dealStatus,
      target: this.target,
      source: this.source,
      includecharge: this.includecharge,
      pack: this.pack,
      toporbottom: this.TotalReport,
      fromdate: this.FromDate,
      todate: this.ToDate,
      GridView: this.GridView,
      groups: this.groups,
      count: 0,
      as: this.acquisition

    };
    this.shared.api.SetHeaderData({ obj: data });
    this.header = [
      {
        type: 'Bar',
        storeIds: this.storeIds,
        fromDate: this.FromDate,
        toDate: this.ToDate,
        ReportTotal: this.TotalReport, groups: this.groups,
        sp: this.salesPersonId,
        sm: this.salesManagerId,
        fm: this.financeManagerId,
        as: this.acquisition, gridview: this.GridView, otherstoreids: this.otherstoreid, selectedotherstoreids: this.selectedotherstoreids,
        datevaluetype: this.DateType, dealType: this.dealType,
        saleType: this.saleType,
        dealStatus: this.dealStatus,

      },
    ];
    if (this.GridView == 'Global') {
      this.getSalesData()
    } else {
      // this.GetBackGrossSalesdata()
    }
    // }
    //     }
    // })
  }

  setDates(type: any) {
    let dates: any = this.setdates.setDates(type)
    this.FromDate = dates[0];
    this.ToDate = dates[1];
    localStorage.setItem('time', type);
  }
  getSalesData() {
    // this.responcestatus = '';
    
    if ((this.storeIds != '' && this.storeIds != '0')) {
      this.shared.spinner.show();
      this.NoData = false;
      this.GetData();
    } else {
      // this.NoData = true;
      this.shared.spinner.hide();
    }

    // this.GetTotalData();
  }

  GetData() {
  
    this.IndividualSalesGross = [];
    // this.shared.spinner.show();
    const obj = {
      startdealdate: this.FromDate,
      enddealdate: this.ToDate,
      StoreID: this.storeIds,
      SalesPerson: this.salesPersonId,
      SalesManager: this.salesManagerId,
      FinanceManager: this.financeManagerId,
      dealtype: this.dealType.toString(),
      saletype: this.saleType.toString(),
      dealstatus: this.dealStatus.toString(),
      AcquisitionSource: this.acquisition.toString() == 'All' ? '' : this.acquisition.toString(),
      var1: this.path1,
      var2: this.path2,
      var3: this.path3,
      Rowtype: 'D',
    };

    const curl = this.shared.getEnviUrl() + this.shared.common.routeEndpoint + 'GetSalesGrossData';
    this.shared.api
      .postmethod(this.shared.common.routeEndpoint + 'GetSalesGrossData', obj)
      .subscribe(
        (res) => {
          const currentTitle = document.title;
          this.shared.api.logSaving(curl, {}, '', res.message, currentTitle);
          if (res.status == 200) {
            this.IndividualSalesGross = [];
            this.TotalSalesGross = [];
            if (res.response != undefined) {
              if (res.response.length > 0) {
                this.IndividualSalesGross = res.response;
                this.TotalSalesGross = res.response.filter(
                  (e: any) => e.data1.toLowerCase() == 'report totals'
                );
                this.TotalSalesGross = this.TotalSalesGross.map((v: any) => ({
                  ...v,
                  Dealer: '+',
                }));
                this.IndividualSalesGross = res.response.filter(
                  (i: any) => i.data1.toLowerCase() != 'report totals'
                );
                // this.responcestatus = 'I';
                this.NoData = false;
                let length = this.IndividualSalesGross.length;
                let path2 = this.path2;
                let path3 = this.path3;
                if (this.TotalReport == 'B') {
                  this.IndividualSalesGross.push(this.TotalSalesGross[0]);
                } else {
                  this.IndividualSalesGross.unshift(this.TotalSalesGross[0]);
                }
                //console.log(this.IndividualSalesGross);

                this.IndividualSalesGross.some(function (x: any) {
                  if (
                    x.data2 != undefined &&
                    x.data2 != '' &&
                    x.data2 != null
                  ) {
                    x.Data2 = JSON.parse(x.data2);
                    x.Data2 = x.Data2.map((v: any) => ({
                      ...v,
                      SubData: [],
                      data2sign: '-',
                    }));
                  }
                  if (
                    x.data3 != undefined &&
                    x.data3 != '' &&
                    x.data3 != null
                  ) {
                    x.Data3 = JSON.parse(x.data3);
                    x.Data2.forEach((val: any) => {
                      x.Data3.forEach((ele: any) => {
                        ele.data3 = ele.data3?.toLowerCase()
                        if (val.data2 == ele.data2) {
                          val.SubData.push(ele);
                        }
                      });
                    });
                  }
                  if (length == 2 || (path2 != '' && path3 == '')) {
                    x.Dealer = '-';
                  } else {
                    x.Dealer = '+';
                  }
                });
                // this.combineIndividualandTotal();

                this.SalesData = this.IndividualSalesGross;

                this.shared.spinner.hide();
                console.log(this.SalesData, '...................');

                // //console.log(this.responcestatus);
                // this.SalesData=this.IndividualSalesGross;
              } else {
                // //this.shared.toast.error('Empty Response','');
                this.shared.spinner.hide();
                this.NoData = true;
                this.SalesData = []
              }
            } else {
             
              //  //this.shared.toast.error('Empty Response', '');
              this.shared.spinner.hide();
              this.NoData = true;
              this.SalesData = []

            }

          } else {
        

            //this.shared.toast.error(res.status, '');
            this.shared.spinner.hide();
            this.NoData = true;
            this.SalesData = []

          }
        },
        (error) => {
          //this.shared.toast.error('502 Bad Gate Way Error', '');
          this.shared.spinner.hide();
          this.NoData = true;
          this.SalesData = []

        }
      );
  }
  GetTotalData() {
    this.TotalSalesGross = [];
    const obj = {
      startdealdate: this.FromDate,
      enddealdate: this.ToDate,
      StoreID: this.storeIds,
      SalesPerson: this.salesPersonId,
      SalesManager: this.salesManagerId,
      FinanceManager: this.financeManagerId,
      dealtype: this.dealType,
      saletype: this.saleType,
      dealstatus: this.dealStatus.toString(),
      var1: this.path1,
      var2: this.path2,
      var3: this.path3,
      Rowtype: 'T',
    };
    this.shared.api
      .postmethod(this.shared.common.routeEndpoint + 'GetSalesGrossData', obj)
      .subscribe(
        (totalres) => {
          if (totalres.status == 200) {
            if (totalres.response != undefined) {
              if (totalres.response.length > 0) {
                this.TotalSalesGross = totalres.response.map((v: any) => ({
                  ...v,
                  Data2: [],
                  Dealer: '+',
                }));
                this.responcestatus = this.responcestatus + 'T';
                this.combineIndividualandTotal();
                // //console.log(this.responcestatus);
              } else {
                // //this.shared.toast.error('Empty Response','');
                this.shared.spinner.hide();
                this.NoData = true;
              }
            } else {
              // //this.shared.toast.error('Empty Response','');
              this.shared.spinner.hide();
              this.NoData = true;
            }
          } else {
            //this.shared.toast.error(totalres.status, '');
            this.shared.spinner.hide();
            this.NoData = true;
          }
        },
        (error) => {
          //this.shared.toast.error('502 Bad Gate Way Error', '');
          this.shared.spinner.hide();
          this.NoData = true;
        }
      );
  }
  combineIndividualandTotal() {
    this.SalesData = this.IndividualSalesGross;
    this.shared.spinner.hide();

    if (this.responcestatus == 'IT' || this.responcestatus == 'TI') {
      if (this.TotalReport == 'B') {
        this.IndividualSalesGross.push(this.TotalSalesGross[0]);
        this.SalesData = this.IndividualSalesGross;
        this.shared.spinner.hide();
        console.log(this.SalesData);
      } else {
        this.IndividualSalesGross.unshift(this.TotalSalesGross[0]);
        this.SalesData = this.IndividualSalesGross;
        this.shared.spinner.hide();
        console.log(this.SalesData);
      }
    } else if (this.responcestatus == 'T') {
      this.SalesData = this.TotalSalesGross;
      this.shared.spinner.hide();
    } else if (this.responcestatus == 'I') {
      this.SalesData = this.IndividualSalesGross;
      this.shared.spinner.hide();
    } else {
      this.NoData = true;
    }
    console.log(this.SalesData, '...................');

  }


  GetBackGrossSalesdata() {
    this.responcestatus = '';
    this.shared.spinner.show();
    this.GridView = 'BackGross';
    this.NoData = false;
    this.getPeopleList()
    this.GetBackGrossData();
    // this.GetBackGrossTotalData();
  }

  GetBackGrossData() {
    this.shared.spinner.show()
    this.GridView = 'BackGross';


    this.IndividualSalesGrossBackGross = [];
    const obj = {
      startdealdate: this.FromDate,
      enddealdate: this.ToDate,
      StoreID: this.otherstoreid && this.selectedotherstoreids ? this.storeIds + ',' + this.selectedotherstoreids.toString() : this.storeIds,
      SalesPerson: this.salesPersonId,
      SalesManager: this.salesManagerId,
      FinanceManager: this.financeManagerId,
      dealtype: this.dealType.toString(),
      saletype: this.saleType.toString(),
      dealstatus: this.dealStatus.toString(),
      var1: this.path1,
      var2: this.path2,
      var3: this.path3,
      Rowtype: 'D',
    };

    this.shared.api
      .postmethod(this.shared.common.routeEndpoint + 'GetSalesGrossSummaryDataBackgross', obj)
      .subscribe(
        (res) => {
          if (res.status == 200) {
            if (res.response != undefined) {
              if (res.response.length > 0) {
                this.IndividualSalesGrossBackGross = [];
                // this.IndividualSalesGrossBackGross = res.response;

                this.IndividualSalesGrossBackGross = res.response;
                this.TotalSalesGrossBackGross = res.response.filter((e: any) => e.data1.toLowerCase() == 'reports total');
                this.TotalSalesGrossBackGross = this.TotalSalesGrossBackGross.map((v: any) => ({ ...v, Dealer: '-', }));
                this.TotalSalesGrossBackGross && this.TotalSalesGrossBackGross.length > 0 ? this.TotalSalesGrossBackGross[0].data1 = 'REPORT TOTALS' : ''
                this.IndividualSalesGrossBackGross = res.response.filter((i: any) => i.data1.toLowerCase() != 'reports total');

                this.responcestatus = this.responcestatus + 'I';
                let length = this.IndividualSalesGrossBackGross.length;
                let path2 = this.path2;
                let path3 = this.path3;

                if (this.TotalReport == 'B') {
                  this.IndividualSalesGrossBackGross.push(this.TotalSalesGrossBackGross[0]);
                } else {
                  this.IndividualSalesGrossBackGross.unshift(this.TotalSalesGrossBackGross[0]);
                }
                this.IndividualSalesGrossBackGross.some(function (x: any) {
                  if (x.data2 != undefined && x.data2 != '') {
                    x.data2 = JSON.parse(x.data2);
                    x.data2 = x.data2.map((v: any) => ({
                      ...v,
                      SubData: [],
                      data2sign: '-',
                    }));
                  }
                  if (x.data3 != undefined && x.data3 != '') {
                    x.data3 = JSON.parse(x.data3);
                    x.data2.forEach((val: any) => {
                      x.data3.forEach((ele: any) => {
                        if (val.data2 == ele.data2) {
                          val.SubData.push(ele);
                        }
                      });
                    });
                  }

                  if (path2 == '' && path3 == '') {
                    x.Dealer = '+';
                  } else if (path2 != '' || path3 == '') {
                    x.Dealer = '-';
                  } else {
                    x.Dealer = '+';
                  }
                });
                this.BackGross = this.IndividualSalesGrossBackGross;
                console.log(this.BackGross);
                this.shared.spinner.hide()
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
          
            this.shared.spinner.hide();
            this.NoData = true;
          }
        },
        (error) => {
          //console.log(error);
          // 
          // this.toast.error('502 Bad Gate Way Error', '');
          this.shared.spinner.hide();
          this.NoData = true;
        }
      );
  }

  GetBackGrossTotalData() {
    this.TotalSalesGrossBackGross = [];
    const obj = {
      startdealdate: this.FromDate,
      enddealdate: this.ToDate,
      StoreID: this.otherstoreid && this.selectedotherstoreids ? this.storeIds + ',' + this.selectedotherstoreids.toString() : this.storeIds,
      SalesPerson: this.salesPersonId,
      SalesManager: this.salesManagerId,
      FinanceManager: this.financeManagerId,
      dealtype: this.dealType.toString(),
      saletype: this.saleType.toString(),
      dealstatus: this.dealStatus.toString(),
      var1: this.path1,
      var2: this.path2,
      var3: this.path3,
      Rowtype: 'T',
    };
    this.shared.api
      .postmethod(this.shared.common.routeEndpoint + 'GetSalesGrossSummaryDataBackgross', obj)
      .subscribe(
        (totalres) => {
          if (totalres.status == 200) {
            if (totalres.response != undefined) {
              if (totalres.response.length > 0) {
                this.TotalSalesGrossBackGross = totalres.response.map(
                  (v: any) => ({
                    ...v,
                    Data2: [],
                    Dealer: '+',
                  })
                );
                this.responcestatus = this.responcestatus + 'T';
                this.combineIndividualandTotalBackGross();
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
        (error) => {
          // this.toast.error('502 Bad Gate Way Error', '');
          this.shared.spinner.hide();
          this.NoData = true;
        }
      );
  }
  combineIndividualandTotalBackGross() {
    if (this.responcestatus == 'IT' || this.responcestatus == 'TI') {
      if (this.TotalReport == 'B') {
        this.IndividualSalesGrossBackGross.push(
          this.TotalSalesGrossBackGross[0]
        );
        this.BackGross = this.IndividualSalesGrossBackGross;
      } else {
        this.IndividualSalesGrossBackGross.unshift(
          this.TotalSalesGrossBackGross[0]
        );
        this.BackGross = this.IndividualSalesGrossBackGross;
      }
    } else if (this.responcestatus == 'T') {
      this.BackGross = this.TotalSalesGrossBackGross;
    } else if (this.responcestatus == 'I') {
      this.BackGross = this.IndividualSalesGrossBackGross;
    } else {
      this.NoData = true;
    }
    this.shared.spinner.hide();
  }

  back2grid() {
    this.GridView = 'Global';
    this.getPeopleList()

    // this.getSalesData();
  }
  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    } else if (value < 0) {
      return false;
    }
    return true;
  }

  openDetails(Item: any, ParentItem: any, cat: any) {
    //console.log(Item, ParentItem, cat);
    if (cat == '3') {
      if (Item.data3 != undefined) {
        const DetailsSalesPeron = this.shared.ngbmodal.open(
          SalesgrossDetails,
          {
            size: 'xl',
            backdrop: 'static',
          }
        );
        DetailsSalesPeron.componentInstance.Salesdetails = [
          {
            StartDate: this.FromDate,
            EndDate: this.ToDate,
            dealtype: this.dealType,
            saletype: this.saleType,
            dealstatus: this.dealStatus,
            var1: this.path1,
            var2: this.path2,
            var3: this.path3,
            var1Value: ParentItem.data1,
            var2Value: Item.data2,
            var3Value: Item.data3,
            userName: Item.data3,
            AcquisitionSrc: this.acquisition,
            SalesPerson: this.salesPersonId,
            SalesManager: this.salesManagerId,
            FinanceManager: this.financeManagerId,
          },
        ];
        DetailsSalesPeron.result.then(
          (data) => { },
          (reason) => {
            // on dismiss
          }
        );
      }
    }
    if (cat == '2') {
      if (Item.data2 != undefined && ParentItem.data1.toLowerCase() != 'report totals') {
        const DetailsSalesPeron = this.shared.ngbmodal.open(
          SalesgrossDetails,
          {
            size: 'xl',
            backdrop: 'static',
          }
        );
        DetailsSalesPeron.componentInstance.Salesdetails = [
          {
            StartDate: this.FromDate,
            EndDate: this.ToDate,
            dealtype: this.dealType,
            saletype: this.saleType,
            dealstatus: this.dealStatus,
            var1: this.path1,
            var2: this.path2,
            var3: '',
            var1Value: ParentItem.data1,
            var2Value: Item.data2,
            var3Value: '',
            userName: Item.data2,
            AcquisitionSrc: this.acquisition,
            SalesPerson: this.salesPersonId,
            SalesManager: this.salesManagerId,
            FinanceManager: this.financeManagerId,

          },
        ];
        DetailsSalesPeron.result.then(
          (data) => { },
          (reason) => {
            // on dismiss
          }
        );
      }
    }
    if (cat == '1') {
      if (Item.data1 != undefined && Item.data1.toLowerCase() != 'report totals') {
        const DetailsSalesPeron = this.shared.ngbmodal.open(
          SalesgrossDetails,
          {
            size: 'xl',
            backdrop: 'static',
          }
        );
        DetailsSalesPeron.componentInstance.Salesdetails = [
          {
            StartDate: this.FromDate,
            EndDate: this.ToDate,
            dealtype: this.dealType,
            saletype: this.saleType,
            dealstatus: this.dealStatus,
            var1: this.path1,
            var2: '',
            var3: '',
            var1Value: Item.data1,
            var2Value: '',
            var3Value: '',
            userName: Item.data1,
            AcquisitionSrc: this.acquisition,
            SalesPerson: this.salesPersonId,
            SalesManager: this.salesManagerId,
            FinanceManager: this.financeManagerId,

          },
        ];
        DetailsSalesPeron.result.then(
          (data) => { },
          (reason) => {
            // on dismiss
          }
        );
      }
    }
  }

  isDesc: boolean = false;
  column: string = 'CategoryName';

  sort(property: any, data: any) {
    this.isDesc = !this.isDesc; //change the direction
    this.column = property;
    let direction = this.isDesc ? 1 : -1;
    // //console.log(direction);
    let individual = data.filter((e: any) => e.data1.toLowerCase() !== 'report totals');
    let total = data.filter((e: any) => e.data1.toLowerCase() == 'report totals');
    console.log(individual, property);

    let sorteddata = individual.sort(function (a: any, b: any) {
      if (a[property] < b[property]) {
        return -1 * direction;
      } else if (a[property] > b[property]) {
        return 1 * direction;
      } else {
        return 0;
      }
    });
    console.log(sorteddata, '................');

    if (this.TotalReport == 'T') {
      return this.SalesData = [...total, ...sorteddata]
    } else {
      return this.SalesData = [...sorteddata, ...total]

    }
  }

  // sort(property: string, data: any[]): any[] {
  //   // toggle sort direction
  //   this.isDesc = !this.isDesc;
  //   this.column = property;

  //   const direction = this.isDesc ? 1 : -1;

  //   // Normalize the label check (handles "Report Total" / "Reports Total")
  //   const isTotalRow = (e: any) =>
  //     (e?.data1 ?? '').toString().trim().toLowerCase() === 'reports total' ||
  //     (e?.data1 ?? '').toString().trim().toLowerCase() === 'report total';

  //   const individual = data.filter(e => !isTotalRow(e));
  //   const total = data.filter(e => isTotalRow(e));

  //   individual.sort((a: any, b: any) => {
  //     const av = a?.[property];
  //     const bv = b?.[property];

  //     // Handle null/undefined
  //     if (av == null && bv == null) return 0;
  //     if (av == null) return -1 * direction;
  //     if (bv == null) return 1 * direction;

  //     // Compare numbers if possible, else compare as strings
  //     const an = Number(av);
  //     const bn = Number(bv);
  //     const bothNumbers = !isNaN(an) && !isNaN(bn);

  //     if (bothNumbers) {
  //       return (an - bn) * direction;
  //     }

  //     // String compare (case-insensitive)
  //     return av.toString().localeCompare(bv.toString(), undefined, { sensitivity: 'base' }) * direction;
  //   });

  //   // If TotalReport === 'T' total first else total last
  //   const result = (this.TotalReport === 'T')
  //     ? [...total, ...individual]
  //     : [...individual, ...total];

  //   return result;
  // }
  subdataindex: any = 0;
  expandorcollapse(ind: any, e: any, ref: any, Item: any, parentData: any) {
    let id = (e.target as Element).id;
    if (id == 'D_' + ind) {
      if (this.path2 == '') {
        this.openDetails(Item, parentData, '1');
      } else {
        if (ref == '-') {
          Item.Dealer = '+';
        }
        if (ref == '+') {
          Item.Dealer = '-';
        }
      }
    }
    if (id == 'DN_' + ind) {
      // // //console.log(this.path3)
      if (this.path3 == '') {
        this.openDetails(Item, parentData, '2');
      } else {
        if (ref == '-') {
          Item.data2sign = '+';
        }
        if (ref == '+') {
          Item.data2sign = '-';
          Item.Dealer = '-';
        }
      }
    }
  }

  ngAfterViewInit() {

    this.shared.api.GetReports().subscribe((data) => {

      // console.log(data);
      if (data.obj.Reference == 'Sales Gross') {
        if (data.obj.header == undefined) {
          // this.TotalReport = data.obj.TotalReport;
          this.DateType = localStorage.getItem('time')
          if (data.obj.FromDate != undefined && data.obj.ToDate != undefined) {
            this.FromDate = data.obj.FromDate;
            this.ToDate = data.obj.ToDate;
            this.storeIds = data.obj.storeValues;
            this.salesPersonId = data.obj.Spvalues;
            this.salesManagerId = data.obj.SMvalues;
            this.financeManagerId = data.obj.FIvalues;
            this.dealType = data.obj.dealType;
            this.saleType = data.obj.saleType;
            this.dealStatus = data.obj.dealStatus;
            this.TotalReport = data.obj.reportTotal
            // this.target = data.obj.target;
            // this.source = data.obj.source;
            // this.includecharge = data.obj;
            // this.pack = data.obj.pack;
            this.path1 =
              data.obj.dataGroupingvalues[0] == undefined
                ? ''
                : data.obj.dataGroupingvalues[0].columnName;
            this.path2 =
              data.obj.dataGroupingvalues[1] == undefined
                ? ''
                : data.obj.dataGroupingvalues[1].columnName;
            this.path3 =
              data.obj.dataGroupingvalues[2] == undefined
                ? ''
                : data.obj.dataGroupingvalues[2].columnName;
            this.path1id =
              data.obj.dataGroupingvalues[0] == undefined
                ? ''
                : data.obj.dataGroupingvalues[0].ARG_ID;
            this.path2id =
              data.obj.dataGroupingvalues[1] == undefined
                ? ''
                : data.obj.dataGroupingvalues[1].ARG_ID;
            this.path3id =
              data.obj.dataGroupingvalues[2] == undefined
                ? ''
                : data.obj.dataGroupingvalues[2].ARG_ID;

            this.path1name =
              data.obj.dataGroupingvalues[0] == undefined
                ? ''
                : data.obj.dataGroupingvalues[0].ARG_LABEL;
            this.path2name =
              data.obj.dataGroupingvalues[1] == undefined
                ? ''
                : data.obj.dataGroupingvalues[1].ARG_LABEL;
            this.path3name =
              data.obj.dataGroupingvalues[2] == undefined
                ? ''
                : data.obj.dataGroupingvalues[2].ARG_LABEL;
            this.groups = data.obj.groups;
            this.DateType = data.obj.dateType
            this.acquisition = data.obj.acquisition;
            this.selectedotherstoreids = data.obj.otherstoreids;
          }
        } else {
          if (data.obj.header == 'Yes') {
            this.storeIds = data.obj.storeValues;
          }
        }
        // console.log(this.GridView);           

        // if (this.GridView == 'Global') {
        // console.log(this.GridView); 
        
        if (this.GridView == 'Global') {
          this.getSalesData();

        } else {
          this.GetBackGrossSalesdata()
        }
        // }

        const headerdata = {
          title: 'Sales Gross',
          path1: this.path1name,
          path2: this.path2name,
          path3: this.path3name,
          path1id: this.path1id,
          path2id: this.path2id,
          path3id: this.path3id,
          stores: this.storeIds,
          salespresons: this.salesPersonId,
          salesmanagers: this.salesManagerId,
          financemanagers: this.financeManagerId,
          dealType: this.dealType,
          saleType: this.saleType,
          dealStatus: this.dealStatus,
          target: this.target,
          source: this.source,
          includecharge: this.includecharge,
          pack: this.pack,
          toporbottom: this.TotalReport,
          fromdate: this.FromDate,
          todate: this.ToDate,
          GridView: this.GridView,
          groups: this.groups,
          as: this.acquisition

        };
        this.shared.api.SetHeaderData({
          obj: headerdata,
        });
        this.header = [
          {
            type: 'Bar',
            storeIds: this.storeIds,
            fromDate: this.FromDate,
            toDate: this.ToDate,
            ReportTotal: this.TotalReport,
            groups: this.groups, sp: this.salesPersonId,
            sm: this.salesManagerId,
            fm: this.financeManagerId,
            as: this.acquisition, gridview: this.GridView, otherstoreids: this.otherstoreid, selectedotherstoreids: this.selectedotherstoreids, datevaluetype: this.DateType, dealType: this.dealType,
            saleType: this.saleType,
            dealStatus: this.dealStatus,

          },
        ];
      }

    });

    this.excel = this.shared.api.getExportToExcelAllReports().subscribe((res) => {
      if (this.excel != undefined) {
        if (res.obj.title == 'Sales Gross') {
          if (res.obj.state == true) {
            this.exportToExcel();
          }
        }
      }
    });


  }


  datetype() {
    if (this.DateType == 'PM') {
      return 'SP';
    }
    else if (this.DateType == 'C') {
      return 'C'
    }
    return this.DateType;
  }


  Scrollpercent: any = 0;
  updateVerticalScroll(event: any): void {
    const scrollDemo = document.querySelector('#scrollcent');
    this.Scrollpercent = Math.round(
      (event.target.scrollTop /
        (event.target.scrollHeight - scrollDemo!.clientHeight)) *
      100
    );
  }

  index = '';
  commentobj = {};

  close(data: any) {
    // //console.log(data);
    this.index = '';
  }
  addcmt(data: any) {
    if (data == 'AD') {
      this.getSalesData();
    }
  }
  ExcelStoreNames: any = []
  exportToExcel(): void {
    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('Sales Gross');

    const PresentYear = this.shared.datePipe.transform(this.FromDate, 'yyyy');
    const FromDate = this.shared.datePipe.transform(this.FromDate, 'dd');
    const ToDate = this.shared.datePipe.transform(this.ToDate, 'dd');
    const PresentMonth = this.shared.datePipe.transform(this.FromDate, 'MMMM');

    let filters: any = [
      { name: 'Groupings :', values: this.path1name + (this.path2name != '' && this.path2name != undefined ? ', ' + this.path2name : '') + (this.path3name != '' && this.path3name != undefined ? ', ' + this.path3name : '') },
      { name: 'Time Frame :', values: this.FromDate + ' to ' + this.ToDate },
      { name: 'Sales Persons :', values: this.salesPersonId == 0 || this.salesPersonId == '' ? 'All Sales Persons' : this.salesPersonId == null ? '-' : this.salesPersonId },
      { name: 'Sales Managers :', values: this.salesManagerId == 0 || this.salesManagerId == '' ? 'All Sales Managers' : this.salesManagerId == null ? '-' : this.salesManagerId },
      { name: 'F&I Managers :', values: this.financeManagerId == 0 || this.financeManagerId == '' ? 'All F&I Managers' : this.financeManagerId == null ? '-' : this.financeManagerId },
      { name: 'New Used : ', values: this.dealType == '' ? '-' : this.dealType == null ? '-' : this.dealType.toString().replaceAll(',', ', ') },
      { name: 'Deal Type :', values: this.saleType == '' ? '-' : this.saleType == null ? '-' : this.saleType.toString().replaceAll(',', ', ').replace('Rental', 'Rental/Loaner') },
      { name: 'Deal Status :', values: this.dealStatus == '' ? '-' : this.dealStatus == null ? '-' : this.dealStatus.toString().replaceAll(',', ', ').replace('Capped', 'Booked').replace('Finalized', 'Finalized') },
    ]
    // const ReportFilter = worksheet.addRow(['Report Controls :']);
    // ReportFilter.font = { name: 'Arial', family: 4, size: 10, bold: true };
    const titleRow = worksheet.addRow(['Sales Gross']);
    titleRow.eachCell((cell, number) => {
      cell.alignment = {
        indent: 1,
        vertical: 'middle',
        horizontal: 'left',
      };
    });
    titleRow.font = { name: 'Arial', family: 4, size: 12, bold: true };
    titleRow.worksheet.mergeCells('A2', 'D2');


    this.ExcelStoreNames = []
    let storeNames: any[] = [];
    let store: any = [];
    this.storeIds && this.storeIds.toString().indexOf(',') > 0 ? store = this.storeIds.toString().split(',') : store.push(this.storeIds)
    console.log(this.storeIds, this.groups, '...............................');

    storeNames = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groups)[0].Stores.filter((item: any) => store.includes(item.ID.toString()));
    if (store.length == this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groups)[0].Stores.length) { this.ExcelStoreNames = 'All Stores' }
    else { this.ExcelStoreNames = storeNames.map(function (a: any) { return a.storename; }); }

    const Stores1 = worksheet.getCell('A3');
    Stores1.value = 'Stores :';
    worksheet.mergeCells('B3', 'Z3');
    const stores1 = worksheet.getCell('B3');
    stores1.value = this.ExcelStoreNames.toString().replaceAll(',', ', ');
    stores1.font = { name: 'Arial', family: 4, size: 9 };
    stores1.alignment = { vertical: 'top', horizontal: 'left', wrapText: true, };
    let startIndex = 3
    filters.forEach((val: any) => {
      startIndex++
      worksheet.addRow('');
      worksheet.getCell(`A${startIndex}`);
      worksheet.mergeCells(`B${startIndex}:C${startIndex}`);
      worksheet.getCell(`A${startIndex}`).value = val.name;
      worksheet.getCell(`B${startIndex}`).value = val.values
    })

    worksheet.addRow('');
    worksheet.getCell('A12');
    worksheet.mergeCells('B12:F12');
    worksheet.mergeCells('G12:K12');
    worksheet.mergeCells('L12:P12');
    worksheet.mergeCells('Q12:U12');

    worksheet.getCell('A12').value = `${PresentMonth}`;
    worksheet.getCell('B12').value = 'UNITS';
    worksheet.getCell('G12').value = 'FRONT GROSS';
    worksheet.getCell('L12').value = 'BACK GROSS';
    worksheet.getCell('Q12').value = 'TOTAL GROSS';


    worksheet.getRow(1).height = 25;


    ['A12', 'B12', 'G12', 'L12', 'Q12'].forEach(key => {
      const cell = worksheet.getCell(key);
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2F5597' } };
    });

    const dateLabel =
      this.datetype() === 'C'
        ? `${this.shared.datePipe.transform(this.FromDate, 'MM.dd.yyyy')}-${this.shared.datePipe.transform(this.ToDate, 'MM.dd.yyyy')}`
        : this.datetype();

    const secondHeader = [
      `${FromDate} - ${ToDate}, ${PresentYear}`,

      dateLabel, 'PACE', 'TARGET', '+/-', 'PER DAY',
      dateLabel, 'PACE', 'TARGET', '+/-', 'PVR',
      dateLabel, 'PACE', 'TARGET', '+/-', 'PVR',
      dateLabel, 'PACE', 'TARGET', '+/-', 'PVR'
    ];

    const headerRow = worksheet.addRow(secondHeader);

    headerRow.eachCell((cell) => {
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.font = { bold: true };
    });

    const bindingHeaders = [
      'data1', 'Units_MTD', 'Units_Pace', 'Units_Target', 'Units_Diff', 'PerDay',
      'FrontGross_MTD', 'FrontGross_Pace', 'FrontGross_Target', 'FrontGross_Diff', 'FrontGross_PVR',
      'BackGross_MTD', 'BackGross_Pace', 'BackGross_Target', 'BackGross_Diff', 'BackGross_PVR',
      'TotalGross_MTD', 'TotalGross_Pace', 'TotalGross_Target', 'TotalGross_Diff', 'TotalGross_PVR',
    ];
    const currencyFields = [
      'FrontGross_MTD', 'FrontGross_Pace', 'FrontGross_Target', 'FrontGross_Diff', 'FrontGross_PVR',
      'BackGross_MTD', 'BackGross_Pace', 'BackGross_Target', 'BackGross_Diff', 'BackGross_PVR',
      'TotalGross_MTD', 'TotalGross_Pace', 'TotalGross_Target', 'TotalGross_Diff', 'TotalGross_PVR',];

    const capitalize = (str: string) =>
      str ? str.toString().replace(/\b\w/g, char => char.toUpperCase()) : '';
    for (const info of this.SalesData) {
      const rowData = bindingHeaders.map(key => {
        const val = info[key];
        if (key === 'data1') return capitalize(val)
        console.log(key, info['FrontGross_Pace'], info['FrontGross_Target'], '.........');

        if (key === 'Units_Diff' && info['Units_Pace'] != 0 && info['Units_Pace'] != null && info['Units_Target'] != 0 && info['Units_Target'] != null) return (info['Units_Pace'] - info['Units_Target']);
        if (key === 'FrontGross_Diff' && info['FrontGross_Pace'] != 0 && info['FrontGross_Pace'] != null && info['FrontGross_Target'] != 0 && info['FrontGross_Target'] != null) return val;
        if (key === 'BackGross_Diff' && info['BackGross_Pace'] != 0 && info['BackGross_Pace'] != null && info['BackGross_Target'] != 0 && info['BackGross_Target'] != null) return val;
        if (key === 'TotalGross_Diff' && info['TotalGross_Pace'] != 0 && info['TotalGross_Pace'] != null && info['TotalGross_Target'] != 0 && info['TotalGross_Target'] != null) return val;

        if (key != 'FrontGross_Diff' && key != 'BackGross_Diff' && key != 'TotalGross_Diff') return val === 0 || val == null ? '-' : val
      });

      const dealerRow = worksheet.addRow(rowData);
      dealerRow.font = { bold: true };

      bindingHeaders.forEach((key, index) => {
        const cell = dealerRow.getCell(index + 1);
        if (currencyFields.includes(key) && typeof cell.value === 'number') {
          cell.numFmt = '"$"#,##0';
          cell.alignment = { horizontal: 'right' };
        } else if (!isNaN(Number(cell.value))) {
          cell.alignment = { horizontal: 'right' };
        }
      });

      if (info.Data2 != undefined) {
        for (const data2 of info.Data2) {
          const nestedRowData = bindingHeaders.map(key => {
            if (key === 'data1') return '   ' + capitalize(data2['data2']);
            if (key === 'FrontGross_PVR') return (data2['FrontGross_MTD'] / data2['Units_MTD']);
            if (key === 'BackGross_PVR') return (data2['BackGross_MTD'] / data2['Units_MTD']);
            if (key === 'TotalGross_PVR') return (data2['TotalGross_MTD'] / data2['Units_MTD']);

            if (key === 'Units_Diff' && data2['Units_Pace'] != 0 && data2['Units_Pace'] != null && data2['Units_Target'] != 0 && data2['Units_Target'] != null) return (data2['Units_Pace'] - data2['Units_Target']);
            if (key === 'FrontGross_Diff' && data2['FrontGross_Pace'] != 0 && data2['FrontGross_Pace'] != null && data2['FrontGross_Target'] != 0 && data2['FrontGross_Target'] != null) return (data2['FrontGross_Pace'] - data2['FrontGross_Target']);
            if (key === 'BackGross_Diff' && data2['BackGross_Pace'] != 0 && data2['BackGross_Pace'] != null && data2['BackGross_Target'] != 0 && data2['BackGross_Target'] != null) return (data2['BackGross_Pace'] - data2['BackGross_Target']);
            if (key === 'TotalGross_Diff' && data2['TotalGross_Pace'] != 0 && data2['TotalGross_Pace'] != null && data2['TotalGross_Target'] != 0 && data2['TotalGross_Target'] != null) return (data2['TotalGross_Pace'] - data2['TotalGross_Target']);

            const val = data2[key];
            return val === 0 || val == null ? '-' : val;
          });
          const nestedRow = worksheet.addRow(nestedRowData);

          bindingHeaders.forEach((key, index) => {
            const cell = nestedRow.getCell(index + 1);
            if (currencyFields.includes(key) && typeof cell.value === 'number') {
              cell.numFmt = '"$"#,##0';
              cell.alignment = { horizontal: 'right' };
            } else if (!isNaN(Number(cell.value))) {
              cell.alignment = { horizontal: 'right' };
            }
          });


          if (data2.SubData != undefined) {
            for (const data3 of data2.SubData) {
              const nestedRowData = bindingHeaders.map(key => {
                if (key === 'data1') return '          ' + capitalize(data3['data3']);
                if (key === 'FrontGross_PVR') return (data3['FrontGross_MTD'] / data3['Units_MTD']);
                if (key === 'BackGross_PVR') return (data3['BackGross_MTD'] / data3['Units_MTD']);
                if (key === 'TotalGross_PVR') return (data3['TotalGross_MTD'] / data3['Units_MTD']);

                if (key === 'Units_Target') return '';
                if (key === 'FrontGross_Target') return '';
                if (key === 'BackGross_Target') return '';
                if (key === 'TotalGross_Target') return '';

                if (key === 'FrontGross_Diff') return '';
                if (key === 'BackGross_Diff') return '';
                if (key === 'TotalGross_Diff') return '';

                const val = data3[key];
                return val === 0 || val == null ? '-' : val;
              });
              const nestedRow = worksheet.addRow(nestedRowData);

              bindingHeaders.forEach((key, index) => {
                const cell = nestedRow.getCell(index + 1);
                if (currencyFields.includes(key) && typeof cell.value === 'number') {
                  cell.numFmt = '"$"#,##0';
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
    worksheet.columns.forEach((column: any) => {
      let maxLength = 20;
      column.width = maxLength + 2;
    });
    workbook.xlsx.writeBuffer().then((buffer: any) => {
      const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      this.shared.exportToExcel(workbook, 'Sales Gross')
    });
  }

  printDiv() {
    const printContents = this.printSection.nativeElement.innerHTML;
    const popupWin = window.open('', '_blank', 'width=800,height=600');

    // Collect all stylesheets from the current document
    const styles = Array.from(document.styleSheets)
      .map((styleSheet: any) => {
        try {
          return Array.from(styleSheet.cssRules)
            .map((rule: any) => rule.cssText)
            .join('');
        } catch (e) {
          return ''; // Ignore CORS-restricted stylesheets
        }
      })
      .join('');

    popupWin!.document.open();
    popupWin!.document.write(`
      <html>
        <head>
          <title>Print Report</title>
          <style>${styles}</style>
        </head>
        <body onload="window.print();window.close()">
          ${printContents}
        </body>
      </html>
    `);
    popupWin!.document.close();
  }
}
