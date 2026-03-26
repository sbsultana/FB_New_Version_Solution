



import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { Component, ElementRef, ViewChild } from '@angular/core';
import { Subscription } from 'rxjs';
import { Setdates } from '../../../../Core/Providers/SetDates/setdates';
import { SalesgrossReports } from '../salesgross-reports/salesgross-reports';
import { SalesgrossDetails } from '../salesgross-details/salesgross-details';
import { CurrencyPipe } from '@angular/common';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
import { NgbActiveModal } from '@ng-bootstrap/ng-bootstrap';
import * as ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';


@Component({
  selector: 'app-dashboard',
  imports: [SharedModule, SalesgrossReports],
  standalone: true,
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss'
})
export class Dashboard {
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

  actionType: any = ''
  DefaultLoad: any = 'E'

  path3id: any = '';
  CurrentDate = new Date();
  NoData: boolean = false;
  responcestatus: string = '';
  groups: any = 1;
  acquisition: any = ['All'];
  otherstoreid: any = '';
  selectedotherstoreids: any = ''

  ProductDeals: any = 'No'
  Months: any = [
    { code: 1, name: 'January' },
    { code: 2, name: 'February' },
    { code: 3, name: 'March' },
    { code: 4, name: 'April' },
    { code: 5, name: 'May' },
    { code: 6, name: 'June' },
    { code: 7, name: 'July' },
    { code: 8, name: 'August' },
    { code: 9, name: 'September' },
    { code: 10, name: 'October' },
    { code: 11, name: 'November' },
    { code: 12, name: 'December' },
  ]

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
      gridview: this.GridView, otherstoreids: this.otherstoreid, selectedotherstoreids: this.selectedotherstoreids, ProductDeals: this.ProductDeals, 'DefaultLoad': this.DefaultLoad
    },
  ];


  reportOpenSub!: Subscription;
  reportGetting!: Subscription;
  Pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;

  selectedDataGrouping: any = [
    { "ARG_ID": 1, "ARG_LABEL": "Store", "ARG_SEQ": 1, "columnname": "store", "Active": 'Y' },
    { "ARG_ID": 2, "ARG_LABEL": "New/Used", "ARG_SEQ": 2, "columnname": "ad_dealtype", "Active": 'Y' },
  ]
  constructor(
    public shared: Sharedservice, public setdates: Setdates, public cp: CurrencyPipe, private toast: ToastService, private ngbmodalActive: NgbActiveModal
  ) {
    this.shared.setTitle(this.shared.common.titleName + '-Sales Gross');
    if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
      if (localStorage.getItem('flag') == 'V') {
        this.storeIds = [];
        console.log(JSON.parse(localStorage.getItem('userInfo')!), JSON.parse(localStorage.getItem('userInfo')!).user_Info, 'Widget Stores............');
        this.groups = JSON.parse(localStorage.getItem('userInfo')!).groupid
        JSON.parse(localStorage.getItem('userInfo')!).store.indexOf(',') > 0 ?
          this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).store.split(',') :
          this.storeIds.push(JSON.parse(localStorage.getItem('userInfo')!).store)
        this.actionType = 'Y';
        this.DefaultLoad = '';
        this.getSalesData();

        localStorage.setItem('flag', 'M')
      } else {
        this.groups = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Preferences
        //this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids.split(',')
        this.storeIds = []
        this.DefaultLoad = 'E'
        this.actionType = 'N'
      }
    }
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
    localStorage.setItem('stime', 'MTD')
    this.getPeopleList()

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
      dataGroupings: this.selectedDataGrouping,
      stores: this.storeIds,
      sp: this.salesPersonId,
      sm: this.salesManagerId,
      fm: this.financeManagerId,
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
      as: this.acquisition,
      datevaluetype: this.DateType,
      'DefaultLoad': this.DefaultLoad

    };
    this.shared.api.SetHeaderData({ obj: data });
    this.header = [
      {
        type: 'Bar',
        dataGroupings: this.selectedDataGrouping,
        stores: this.storeIds,
        sp: this.salesPersonId,
        sm: this.salesManagerId,
        fm: this.financeManagerId,
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
        as: this.acquisition,
        datevaluetype: this.DateType,
        ProductDeals: this.ProductDeals,
        'DefaultLoad': this.DefaultLoad
      },
    ];
    // this.getSalesData()
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
    if ((this.storeIds != '' && this.storeIds != '0') || this.selectedotherstoreids != '') {
      this.shared.spinner.show();
      this.NoData = false;
      this.actionType = 'Y';
      this.DefaultLoad = ''
      if (this.ProductDeals == 'No') {
        this.GetData();
      } else {
        this.GetDataDeals()
      }

    } else {
      // this.NoData = true;
      this.shared.spinner.hide();
    }

    // this.GetTotalData();
  }
  GetBackGrossSalesdata() {
    this.responcestatus = '';

    this.shared.spinner.show();
    this.GetBackGrossData();
    // this.GetBackGrossTotalData();
  }
  GetData() {
    this.IndividualSalesGross = [];
    // this.shared.spinner.show();
    const obj = {
      startdealdate: this.FromDate,
      enddealdate: this.ToDate,
      StoreID: this.selectedotherstoreids != undefined && this.selectedotherstoreids != '' && this.selectedotherstoreids != null ?
        (this.storeIds != '' ? this.storeIds + ',' + this.selectedotherstoreids.toString() : this.selectedotherstoreids.toString()) : this.storeIds,
      SalesPerson: this.salesPersonId,
      SalesManager: this.salesManagerId,
      FinanceManager: this.financeManagerId,
      dealtype: this.dealType.toString(),
      saletype: this.saleType.toString(),
      dealstatus: this.dealStatus.toString(),
      AcquisitionSource: this.acquisition.toString() == 'All' ? '' : this.acquisition.toString(),
      var1: this.selectedDataGrouping.length >= 1 ? this.selectedDataGrouping[0]?.columnname : '',
      var2: this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '',
      var3: this.selectedDataGrouping.length == 3 ? this.selectedDataGrouping[2]?.columnname : '',
      Rowtype: 'D',
    };

    this.shared.api
      .postmethod(this.shared.common.routeEndpoint + 'GetSalesGrossData', obj)
      .subscribe(
        (res) => {
          const currentTitle = document.title;
          if (res.status == 200) {
            this.IndividualSalesGross = [];
            this.TotalSalesGross = [];
            if (res.response != undefined) {
              if (res.response.length > 0) {
                const monthCodeByName = new Map(
                  this.Months.map((m: any) => [m.name.toLowerCase(), m.code])
                );
                if (this.selectedDataGrouping[0]?.columnname == 'ad_Month') {
                  this.IndividualSalesGross = res.response.filter(
                    (e: any) => e.data1 != 'REPORTS TOTAL'
                  ).map((item: any) => {
                    const key = String(item.data1 ?? '').trim().toLowerCase();
                    if (!monthCodeByName.has(key)) {
                      throw new Error(`Unknown month in data1: '${item.data1}'`);
                    }
                    return { ...item, code: monthCodeByName.get(key)! };
                  }).sort((a: any, b: any) => a.code - b.code);
                } else {
                  this.IndividualSalesGross = res.response.filter(
                    (e: any) => e.data1 != 'REPORTS TOTAL'
                  )
                }
                this.TotalSalesGross = res.response.filter(
                  (e: any) => e.data1 == 'REPORTS TOTAL'
                );
                this.TotalSalesGross = this.TotalSalesGross.map((v: any) => ({
                  ...v,
                  Dealer: '+',
                }));
                // this.IndividualSalesGross = res.response.filter(
                //   (i: any) => i.data1 != 'REPORTS TOTAL'
                // );
                // this.responcestatus = 'I';
                this.NoData = false;
                let length = this.IndividualSalesGross.length;
                let path2 = this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '';
                let path3 = this.selectedDataGrouping.length >= 3 ? this.selectedDataGrouping[2]?.columnname : '';
                if (this.TotalReport == 'B') {
                  this.IndividualSalesGross.push(this.TotalSalesGross[0]);
                } else {
                  this.IndividualSalesGross.unshift(this.TotalSalesGross[0]);
                }

                this.IndividualSalesGross.some(function (x: any) {
                  if (
                    x.data2 != undefined &&
                    x.data2 != '' &&
                    x.data2 != null
                  ) {
                    if (path2 == 'ad_Month') {
                      x.Data2 = JSON.parse(x.data2);
                      x.Data2 = x.Data2.map((item: any) => {
                        const key = String(item.data2 ?? '').trim().toLowerCase();
                        if (!monthCodeByName.has(key)) {
                          throw new Error(`Unknown month in data1: '${item.data2}'`);
                        }
                        return {
                          ...item, SubData: [],
                          data2sign: '-', code: monthCodeByName.get(key)!
                        };
                      }).sort((a: any, b: any) => a.code - b.code);
                    } else {
                      x.Data2 = JSON.parse(x.data2);
                      x.Data2 = x.Data2.map((v: any) => ({
                        ...v,
                        SubData: [],
                        data2sign: '-',
                      }));
                    }

                  }
                  if (x.data3 != undefined && x.data3 != '' && x.data3 != null) {

                    if (path3 == 'ad_Month') {
                      x.Data3 = JSON.parse(x.data3);
                      x.Data3 = x.Data3.map((item: any) => {
                        const key = String(item.data3 ?? '').trim().toLowerCase();
                        if (!monthCodeByName.has(key)) {
                          throw new Error(`Unknown month in data1: '${item.data3}'`);
                        }
                        return {
                          ...item, SubData: [],
                          data2sign: '-', code: monthCodeByName.get(key)!
                        };
                      }).sort((a: any, b: any) => a.code - b.code);

                      x.Data2.forEach((val: any) => {
                        x.Data3.forEach((ele: any) => {
                          if (val.data2 == ele.data2) {
                            val.SubData.push(ele);
                          }
                        });
                      });

                    } else {
                      x.Data3 = JSON.parse(x.data3);
                      x.Data2.forEach((val: any) => {
                        x.Data3.forEach((ele: any) => {
                          if (val.data2 == ele.data2) {
                            val.SubData.push(ele);
                          }
                        });
                      });
                    }
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

              } else {
                this.shared.spinner.hide();
                this.NoData = true;
                this.SalesData = []
              }
            } else {
              this.shared.spinner.hide();
              this.NoData = true;
              this.SalesData = []

            }

          } else {
            // alert(res.status);

            this.toast.show(res.status, 'danger', 'Error');

            this.shared.spinner.hide();
            this.NoData = true;
            this.SalesData = []

          }
        },
        (error) => {
          this.toast.show('502 Bad Gate Way Error', 'danger', 'Error');
          this.shared.spinner.hide();
          this.NoData = true;
          this.SalesData = []

        }
      );
  }


  GetDataDeals() {
    this.IndividualSalesGross = [];
    // this.shared.spinner.show();
    const obj = {
      startdealdate: this.FromDate,
      enddealdate: this.ToDate,
      StoreID: this.selectedotherstoreids != undefined && this.selectedotherstoreids != '' && this.selectedotherstoreids != null ?
        (this.storeIds != '' ? this.storeIds + ',' + this.selectedotherstoreids.toString() : this.selectedotherstoreids.toString()) : this.storeIds,
      SalesPerson: this.salesPersonId,
      SalesManager: this.salesManagerId,
      FinanceManager: this.financeManagerId,
      dealtype: this.dealType.toString(),
      saletype: this.saleType.toString(),
      dealstatus: this.dealStatus.toString(),
      AcquisitionSource: this.acquisition.toString() == 'All' ? '' : this.acquisition.toString(),
      var1: this.selectedDataGrouping.length >= 1 ? this.selectedDataGrouping[0]?.columnname : '',
      var2: this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '',
      var3: this.selectedDataGrouping.length == 3 ? this.selectedDataGrouping[2]?.columnname : '',
      Rowtype: 'D',
    };

    this.shared.api
      .postmethod(this.shared.common.routeEndpoint + 'GetSalesGrossDatabyProducts', obj)
      .subscribe(
        (res) => {
          const currentTitle = document.title;
          if (res.status == 200) {
            this.IndividualSalesGross = [];
            this.TotalSalesGross = [];
            if (res.response != undefined) {
              if (res.response.length > 0) {
                const monthCodeByName = new Map(
                  this.Months.map((m: any) => [m.name.toLowerCase(), m.code])
                );
                if (this.selectedDataGrouping[0]?.columnname == 'ad_Month') {
                  this.IndividualSalesGross = res.response.filter(
                    (e: any) => e.data1 != 'REPORTS TOTAL'
                  ).map((item: any) => {
                    const key = String(item.data1 ?? '').trim().toLowerCase();
                    if (!monthCodeByName.has(key)) {
                      throw new Error(`Unknown month in data1: '${item.data1}'`);
                    }
                    return { ...item, code: monthCodeByName.get(key)! };
                  }).sort((a: any, b: any) => a.code - b.code);
                } else {
                  this.IndividualSalesGross = res.response.filter(
                    (e: any) => e.data1 != 'REPORTS TOTAL'
                  )
                }
                this.TotalSalesGross = res.response.filter(
                  (e: any) => e.data1 == 'REPORTS TOTAL'
                );
                this.TotalSalesGross = this.TotalSalesGross.map((v: any) => ({
                  ...v,
                  Dealer: '+',
                }));
                // this.IndividualSalesGross = res.response.filter(
                //   (i: any) => i.data1 != 'REPORTS TOTAL'
                // );
                // this.responcestatus = 'I';
                this.NoData = false;
                let length = this.IndividualSalesGross.length;
                let path2 = this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '';
                let path3 = this.selectedDataGrouping.length >= 3 ? this.selectedDataGrouping[2]?.columnname : '';
                if (this.TotalReport == 'B') {
                  this.IndividualSalesGross.push(this.TotalSalesGross[0]);
                } else {
                  this.IndividualSalesGross.unshift(this.TotalSalesGross[0]);
                }

                this.IndividualSalesGross.some(function (x: any) {
                  if (
                    x.data2 != undefined &&
                    x.data2 != '' &&
                    x.data2 != null
                  ) {
                    if (path2 == 'ad_Month') {
                      x.Data2 = JSON.parse(x.data2);
                      x.Data2 = x.Data2.map((item: any) => {
                        const key = String(item.data2 ?? '').trim().toLowerCase();
                        if (!monthCodeByName.has(key)) {
                          throw new Error(`Unknown month in data1: '${item.data2}'`);
                        }
                        return {
                          ...item, SubData: [],
                          data2sign: '-', code: monthCodeByName.get(key)!
                        };
                      }).sort((a: any, b: any) => a.code - b.code);
                    } else {
                      x.Data2 = JSON.parse(x.data2);
                      x.Data2 = x.Data2.map((v: any) => ({
                        ...v,
                        SubData: [],
                        data2sign: '-',
                      }));
                    }

                  }
                  if (x.data3 != undefined && x.data3 != '' && x.data3 != null) {

                    if (path3 == 'ad_Month') {
                      x.Data3 = JSON.parse(x.data3);
                      x.Data3 = x.Data3.map((item: any) => {
                        const key = String(item.data3 ?? '').trim().toLowerCase();
                        if (!monthCodeByName.has(key)) {
                          throw new Error(`Unknown month in data1: '${item.data3}'`);
                        }
                        return {
                          ...item, SubData: [],
                          data2sign: '-', code: monthCodeByName.get(key)!
                        };
                      }).sort((a: any, b: any) => a.code - b.code);

                      x.Data2.forEach((val: any) => {
                        x.Data3.forEach((ele: any) => {
                          if (val.data2 == ele.data2) {
                            val.SubData.push(ele);
                          }
                        });
                      });

                    } else {
                      x.Data3 = JSON.parse(x.data3);
                      x.Data2.forEach((val: any) => {
                        x.Data3.forEach((ele: any) => {
                          if (val.data2 == ele.data2) {
                            val.SubData.push(ele);
                          }
                        });
                      });
                    }
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


              } else {
                this.shared.spinner.hide();
                this.NoData = true;
                this.SalesData = []
              }
            } else {
              this.shared.spinner.hide();
              this.NoData = true;
              this.SalesData = []

            }

          } else {
            // alert(res.status);

            this.toast.show(res.status, 'danger', 'Error');

            this.shared.spinner.hide();
            this.NoData = true;
            this.SalesData = []

          }
        },
        (error) => {
          this.toast.show('502 Bad Gate Way Error', 'danger', 'Error');
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
      var1: this.selectedDataGrouping.length >= 1 ? this.selectedDataGrouping[0]?.columnname : '',
      var2: this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '',
      var3: this.selectedDataGrouping.length == 3 ? this.selectedDataGrouping[2]?.columnname : '',
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
              } else {
                this.shared.spinner.hide();
                this.NoData = true;
              }
            } else {
              this.shared.spinner.hide();
              this.NoData = true;
            }
          } else {
            this.toast.show(totalres.status, 'danger', 'Error');
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
    if (cat == '3') {
      if (Item.data3 != undefined) {
        const DetailsSalesPeron = this.shared.ngbmodal.open(
          SalesgrossDetails,
          {
            size: 'xxl',
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
            var1: this.selectedDataGrouping.length >= 1 ? this.selectedDataGrouping[0]?.columnname : '',
            var2: this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '',
            var3: this.selectedDataGrouping.length == 3 ? this.selectedDataGrouping[2]?.columnname : '',
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
      if (Item.data2 != undefined && ParentItem.data1 != 'REPORTS TOTAL') {
        const DetailsSalesPeron = this.shared.ngbmodal.open(
          SalesgrossDetails,
          {
            size: 'xxl',
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
            var1: this.selectedDataGrouping.length >= 1 ? this.selectedDataGrouping[0]?.columnname : '',
            var2: this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '',
            var3: this.selectedDataGrouping.length == 3 ? this.selectedDataGrouping[2]?.columnname : '',
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
      if (Item.data1 != undefined && Item.data1 != 'Reports Total') {
        const DetailsSalesPeron = this.shared.ngbmodal.open(
          SalesgrossDetails,
          {
            // size:'xl',
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
            var1: this.selectedDataGrouping.length >= 1 ? this.selectedDataGrouping[0]?.columnname : '',
            var2: this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '',
            var3: this.selectedDataGrouping.length == 3 ? this.selectedDataGrouping[2]?.columnname : '',
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
    if (direction == -1) {
      this.TotalSortPosition = 'T';
    } else {
      this.TotalSortPosition = 'B';
    }
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
  subdataindex: any = 0;
  expandorcollapse(ind: any, e: any, ref: any, Item: any, parentData: any) {
    let id = (e.target as Element).id;
    if (id == 'D_' + ind) {
      if (this.selectedDataGrouping[1]?.columnname == '') {
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
      if (this.selectedDataGrouping[2]?.columnname == '') {
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
    this.reportOpenSub = this.shared.api.GetReportOpening().subscribe((res) => {
      if (this.reportOpenSub != undefined) {
        if (res.obj.Module == 'Sales Gross') {
          document.getElementById('report')?.click();
        }
      }
    });
    this.excel = this.shared.api.getExportToExcelAllReports().subscribe((res) => {
      if (this.excel != undefined) {
        if (res.obj.title == 'Sales Gross') {
          if (res.obj.state == true) {
            if (this.GridView == 'Global') {
              this.exportToExcelSalesGross();
            } else {
              this.exportToExcelBackGross();
            }
          }
        }
      }
    });
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res) => {
      if (this.email != undefined) {
        if (res.obj.title == 'Sales Gross') {
          if (res.obj.stateEmailPdf == true) {
            // alert('HI')
            // this.sendEmailData(res.obj.Email, res.obj.notes, res.obj.from);
          }
        }
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res) => {
      if (this.print != undefined) {
        if (res.obj.title == 'Sales Gross') {
          if (res.obj.statePrint == true) {
            // this.GetPrintData();
          }
        }
      }
    });

    this.Pdf = this.shared.api.getExportToPDFAllReports().subscribe((res) => {
      if (this.Pdf != undefined) {

        if (res.obj.title == 'Sales Gross') {
          if (res.obj.statePDF == true) {
            // this.generatePDF();
            // this.Pdf.unsubscribe()
          }
        }
      }
    });
    this.reportGetting = this.shared.api.GetReports().subscribe((data) => {
      if (this.reportGetting != undefined) {
        if (data.obj.Reference == 'Sales Gross') {
          this.TotalReport = data.obj.reportTotal;
          this.actionType = 'Y'
          this.DateType = localStorage.getItem('time')
          this.FromDate = data.obj.FromDate;
          this.ToDate = data.obj.ToDate;
          this.storeIds = data.obj.storeValues;
          this.salesPersonId = data.obj.Spvalues;
          this.salesManagerId = data.obj.SMvalues;
          this.financeManagerId = data.obj.FIvalues;
          this.dealType = data.obj.dealType;
          this.saleType = data.obj.saleType;
          this.dealStatus = data.obj.dealStatus;
          this.target = data.obj.target;
          this.source = data.obj.source;
          this.includecharge = data.obj;
          this.pack = data.obj.pack;
          this.selectedDataGrouping = data.obj.dataGroupingvalues;
          this.groups = data.obj.groups;
          this.acquisition = data.obj.acquisition;
          this.ProductDeals = data.obj.ProductDeals;
          this.selectedotherstoreids = data.obj.otherstoreids;
          if (this.GridView == 'Global') {
            this.getSalesData();
          } else {
            this.GetBackGrossSalesdata();
          }
          this.getPeopleList()
        }
      }
    });
  }
  ngOnDestroy() {


    if (this.reportOpenSub != undefined) {
      this.reportOpenSub.unsubscribe()
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

  datetype() {
    if (this.DateType == 'PM') {
      return 'SP';
    }
    else if (this.DateType == 'C') {
      return 'C'
    }
    return this.DateType;
  }


  GetBackGrossData() {
    this.shared.spinner.show()
    this.GridView = 'BackGross';
    this.getPeopleList()

    this.IndividualSalesGrossBackGross = [];
    const obj = {
      startdealdate: this.FromDate,
      enddealdate: this.ToDate,
      StoreID: this.selectedotherstoreids != undefined && this.selectedotherstoreids != '' && this.selectedotherstoreids != null ?
        (this.storeIds != '' ? this.storeIds + ',' + this.selectedotherstoreids.toString() : this.selectedotherstoreids.toString()) : this.storeIds,
      SalesPerson: this.salesPersonId,
      SalesManager: this.salesManagerId,
      FinanceManager: this.financeManagerId,
      dealtype: this.dealType.toString(),
      saletype: this.saleType.toString(),
      dealstatus: this.dealStatus.toString(),
      var1: this.selectedDataGrouping.length >= 1 ? this.selectedDataGrouping[0]?.columnname : '',
      var2: this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '',
      var3: this.selectedDataGrouping.length == 3 ? this.selectedDataGrouping[2]?.columnname : '',
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
                const monthCodeByName = new Map(
                  this.Months.map((m: any) => [m.name.toLowerCase(), m.code])
                );
                if (this.selectedDataGrouping[0]?.columnname == 'ad_Month') {
                  this.IndividualSalesGross = res.response.filter(
                    (e: any) => e.data1 != 'REPORTS TOTAL'
                  ).map((item: any) => {
                    const key = String(item.data1 ?? '').trim().toLowerCase();
                    if (!monthCodeByName.has(key)) {
                      throw new Error(`Unknown month in data1: '${item.data1}'`);
                    }
                    return { ...item, code: monthCodeByName.get(key)! };
                  }).sort((a: any, b: any) => a.code - b.code);
                } else {
                  this.IndividualSalesGross = res.response.filter(
                    (e: any) => e.data1 != 'REPORTS TOTAL'
                  )
                }
                // this.IndividualSalesGrossBackGross = res.response;

                // this.IndividualSalesGrossBackGross = res.response;
                this.TotalSalesGrossBackGross = res.response.filter((e: any) => e.data1 == 'REPORTS TOTAL');
                this.TotalSalesGrossBackGross = this.TotalSalesGrossBackGross.map((v: any) => ({ ...v, Dealer: '-', }));
                this.IndividualSalesGrossBackGross = res.response.filter((i: any) => i.data1 != 'REPORTS TOTAL');

                this.responcestatus = this.responcestatus + 'I';
                let length = this.IndividualSalesGrossBackGross.length;
                let path2 = this.selectedDataGrouping.length >= 2 ? this.selectedDataGrouping[1]?.columnname : '';
                let path3 = this.selectedDataGrouping.length >= 3 ? this.selectedDataGrouping[2]?.columnname : '';

                if (this.TotalReport == 'B') {
                  this.IndividualSalesGrossBackGross.push(this.TotalSalesGrossBackGross[0]);
                } else {
                  this.IndividualSalesGrossBackGross.unshift(this.TotalSalesGrossBackGross[0]);
                }
                this.IndividualSalesGrossBackGross.some(function (x: any) {
                  if (x.data2 != undefined && x.data2 != '') {
                    if (path2 == 'ad_Month') {
                      x.data2 = JSON.parse(x.data2);
                      x.data2 = x.data2.map((item: any) => {
                        const key = String(item.data2 ?? '').trim().toLowerCase();
                        if (!monthCodeByName.has(key)) {
                          throw new Error(`Unknown month in data1: '${item.data2}'`);
                        }
                        return {
                          ...item, SubData: [],
                          data2sign: '-', code: monthCodeByName.get(key)!
                        };
                      }).sort((a: any, b: any) => a.code - b.code);
                    } else {
                      x.data2 = JSON.parse(x.data2);
                      x.data2 = x.data2.map((v: any) => ({
                        ...v,
                        SubData: [],
                        data2sign: '-',
                      }));
                    }
                  }
                  if (x.data3 != undefined && x.data3 != '') {
                    if (path3 == 'ad_Month') {
                      x.data3 = JSON.parse(x.data3);
                      x.data3 = x.data3.map((item: any) => {
                        const key = String(item.data3 ?? '').trim().toLowerCase();
                        if (!monthCodeByName.has(key)) {
                          throw new Error(`Unknown month in data1: '${item.data3}'`);
                        }
                        return {
                          ...item, SubData: [],
                          data2sign: '-', code: monthCodeByName.get(key)!
                        };
                      }).sort((a: any, b: any) => a.code - b.code);

                      x.data2.forEach((val: any) => {
                        x.data3.forEach((ele: any) => {
                          if (val.data2 == ele.data2) {
                            val.SubData.push(ele);
                          }
                        });
                      });

                    } else {
                      x.data3 = JSON.parse(x.data3);
                      x.data2.forEach((val: any) => {
                        x.data3.forEach((ele: any) => {
                          if (val.data2 == ele.data2) {
                            val.SubData.push(ele);
                          }
                        });
                      });
                    }
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
                // this.toast.show('Empty Response','');
                this.shared.spinner.hide();
                this.NoData = true;
              }
            } else {
              // this.toast.show('Empty Response','');
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



  back2grid() {
    this.GridView = 'Global';
    this.getPeopleList()
    this.getSalesData();
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


  screenheight: any = 0; divheight: any = 0; trposition: any = 0;

  ExcelStoreNames: any = [];


  // exportToExcel(): void {
  //   const workbook = this.shared.getWorkbook();
  //   const worksheet = workbook.addWorksheet('Sales Gross');
  //   let storeNames: any[] = [];
  //   const store = this.storeIds
  //   storeNames = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groups)[0].Stores.filter((item: any) => store.includes(item.ID));
  //   if (store.length == this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groups)[0].Stores.length) { this.ExcelStoreNames = 'All Stores' }
  //   else { this.ExcelStoreNames = storeNames.map(function (a: any) { return a.storename; }); }
  //   const PresentYear = this.shared.datePipe.transform(this.FromDate, 'yyyy');
  //   const FromDate = this.shared.datePipe.transform(this.FromDate, 'dd');
  //   const ToDate = this.shared.datePipe.transform(this.ToDate, 'dd');
  //   const PresentMonth = this.shared.datePipe.transform(this.FromDate, 'MMMM');

  //   let filters: any = [
  //     // { name: 'Stores :', values: this.ExcelStoreNames.toString() },
  //     { name: 'Groupings :', values: this.selectedDataGrouping[0]?.ARG_LABEL + (this.selectedDataGrouping[1]?.ARG_LABEL != '' && this.selectedDataGrouping[1]?.ARG_LABEL != undefined ? ', ' + this.selectedDataGrouping[1]?.ARG_LABEL : '') + (this.selectedDataGrouping[2]?.ARG_LABEL != '' && this.selectedDataGrouping[2]?.ARG_LABEL != undefined ? ', ' + this.selectedDataGrouping[2]?.ARG_LABEL : '') },
  //     { name: 'Time Frame :', values: this.FromDate + ' to ' + this.ToDate },
  //     { name: 'Sales Persons :', values: this.salesPersonId == 0 || this.salesPersonId == '' ? 'All Sales Persons' : this.salesPersonId == null ? '-' : this.salesPersonId },
  //     { name: 'Sales Managers :', values: this.salesManagerId == 0 || this.salesManagerId == '' ? 'All Sales Managers' : this.salesManagerId == null ? '-' : this.salesManagerId },
  //     { name: 'F&I Managers :', values: this.financeManagerId == 0 || this.financeManagerId == '' ? 'All F&I Managers' : this.financeManagerId == null ? '-' : this.financeManagerId },
  //     { name: 'New Used : ', values: this.dealType == '' ? '-' : this.dealType == null ? '-' : this.dealType.toString().replaceAll(',', ', ') },
  //     { name: 'Deal Type :', values: this.saleType == '' ? '-' : this.saleType == null ? '-' : this.saleType.toString().replaceAll(',', ', ') },
  //     { name: 'Deal Status :', values: this.dealStatus == '' ? '-' : this.dealStatus == null ? '-' : this.dealStatus.toString().replaceAll(',', ', ').replace('Capped', 'Booked') },
  //   ]
  //   // const ReportFilter = worksheet.addRow(['Report Controls :']);
  //   // ReportFilter.font = { name: 'Arial', family: 4, size: 10, bold: true };
  //   const titleRow = worksheet.addRow(['Sales Gross']);
  //   titleRow.eachCell((cell, number) => {
  //     cell.alignment = {
  //       indent: 1,
  //       vertical: 'middle',
  //       horizontal: 'left',
  //     };
  //   });
  //   titleRow.font = { name: 'Arial', family: 4, size: 12, bold: true };
  //   titleRow.worksheet.mergeCells('A2', 'D2');

  //   const Stores1 = worksheet.getCell('A3');
  //   Stores1.value = 'Stores :';
  //   worksheet.mergeCells('B3', 'Z3');
  //   const stores1 = worksheet.getCell('B3');
  //   stores1.value = this.ExcelStoreNames.toString().replaceAll(',', ', ');
  //   stores1.font = { name: 'Arial', family: 4, size: 9 };
  //   stores1.alignment = { vertical: 'top', horizontal: 'left', wrapText: true, };


  //   let startIndex = 3
  //   filters.forEach((val: any) => {
  //     startIndex++
  //     worksheet.addRow('');
  //     worksheet.getCell(`A${startIndex}`);
  //     worksheet.mergeCells(`B${startIndex}:C${startIndex}`);
  //     worksheet.getCell(`A${startIndex}`).value = val.name;
  //     worksheet.getCell(`B${startIndex}`).value = val.values
  //   })

  //   worksheet.addRow('');
  //   worksheet.getCell('A12');
  //   worksheet.mergeCells('B12:F12');
  //   worksheet.mergeCells('G12:K12');
  //   worksheet.mergeCells('L12:P12');
  //   worksheet.mergeCells('Q12:U12');

  //   worksheet.getCell('A12').value = `${PresentMonth}`;
  //   worksheet.getCell('B12').value = 'UNITS';
  //   worksheet.getCell('G12').value = 'FRONT GROSS';
  //   worksheet.getCell('L12').value = 'BACK GROSS';
  //   worksheet.getCell('Q12').value = 'TOTAL GROSS';


  //   worksheet.getRow(1).height = 25;


  //   ['A12', 'B12', 'G12', 'L12', 'Q12'].forEach(key => {
  //     const cell = worksheet.getCell(key);
  //     cell.alignment = { vertical: 'middle', horizontal: 'center' };
  //     cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  //     cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2F5597' } };
  //   });



  //   const dateLabel =
  //     this.datetype() === 'C'
  //       ? `${this.shared.datePipe.transform(this.FromDate, 'MM.dd.yyyy')}-${this.shared.datePipe.transform(this.ToDate, 'MM.dd.yyyy')}`
  //       : this.datetype();

  //   const secondHeader = [
  //     `${FromDate} - ${ToDate}, ${PresentYear}`,

  //     dateLabel, 'PACE', 'TARGET', '+/-', 'PER DAY',
  //     dateLabel, 'PACE', 'TARGET', '+/-', 'PVR',
  //     dateLabel, 'PACE', 'TARGET', '+/-', 'PVR',
  //     dateLabel, 'PACE', 'TARGET', '+/-', 'PVR'
  //   ];

  //   const headerRow = worksheet.addRow(secondHeader);

  //   headerRow.eachCell((cell) => {
  //     cell.alignment = { horizontal: 'center', vertical: 'middle' };
  //     cell.font = { bold: true };
  //   });

  //   const bindingHeaders = [
  //     'data1', 'Units_MTD', 'Units_Pace', 'Units_Target', 'Units_Diff', 'PerDay',
  //     'FrontGross_MTD', 'FrontGross_Pace', 'FrontGross_Target', 'FrontGross_Diff', 'FrontGross_PVR',
  //     'BackGross_MTD', 'BackGross_Pace', 'BackGross_Target', 'BackGross_Diff', 'BackGross_PVR',
  //     'TotalGross_MTD', 'TotalGross_Pace', 'TotalGross_Target', 'TotalGross_Diff', 'TotalGross_PVR',
  //   ];
  //   const currencyFields = [
  //     'FrontGross_MTD', 'FrontGross_Pace', 'FrontGross_Target', 'FrontGross_Diff', 'FrontGross_PVR',
  //     'BackGross_MTD', 'BackGross_Pace', 'BackGross_Target', 'BackGross_Diff', 'BackGross_PVR',
  //     'TotalGross_MTD', 'TotalGross_Pace', 'TotalGross_Target', 'TotalGross_Diff', 'TotalGross_PVR',];

  //   const capitalize = (str: string) =>
  //     str ? str.toString().replace(/\b\w/g, char => char.toUpperCase()) : '';

  //   for (const info of this.SalesData) {
  //     const rowData = bindingHeaders.map(key => {
  //       const val = info[key];
  //       if (key === 'data1') return capitalize(val)
  //       if (key === 'Units_Diff' && info['Units_Pace'] != 0 && info['Units_Pace'] != null && info['Units_Target'] != 0 && info['Units_Target'] != null) return (info['Units_Pace'] - info['Units_Target']);
  //       if (key === 'FrontGross_Diff' && info['FrontGross_Pace'] != 0 && info['FrontGross_Pace'] != null && info['FrontGross_Target'] != 0 && info['FrontGross_Target'] != null) return val;
  //       if (key === 'BackGross_Diff' && info['BackGross_Pace'] != 0 && info['BackGross_Pace'] != null && info['BackGross_Target'] != 0 && info['BackGross_Target'] != null) return val;
  //       if (key === 'TotalGross_Diff' && info['TotalGross_Pace'] != 0 && info['TotalGross_Pace'] != null && info['TotalGross_Target'] != 0 && info['TotalGross_Target'] != null) return val;

  //       if (key != 'FrontGross_Diff' && key != 'BackGross_Diff' && key != 'TotalGross_Diff') return val === 0 || val == null ? '-' : val
  //     });

  //     const dealerRow = worksheet.addRow(rowData);
  //     dealerRow.font = { bold: true };

  //     bindingHeaders.forEach((key, index) => {
  //       const cell = dealerRow.getCell(index + 1);
  //       if (currencyFields.includes(key) && typeof cell.value === 'number') {
  //         cell.numFmt = '"$"#,##0';
  //         cell.alignment = { horizontal: 'right' };
  //         if (cell.value < 0) {
  //           cell.font = { color: { argb: 'FFFF0000' }, }
  //         }
  //       } else if (!isNaN(Number(cell.value))) {
  //         cell.alignment = { horizontal: 'right' };
  //       }
  //     });

  //     if (info.Data2 != undefined) {
  //       for (const data2 of info.Data2) {
  //         const nestedRowData = bindingHeaders.map(key => {
  //           if (key === 'data1') return '   ' + capitalize(data2['data2']);
  //           if (key === 'FrontGross_PVR') return (data2['FrontGross_MTD'] / data2['Units_MTD']);
  //           if (key === 'BackGross_PVR') return (data2['BackGross_MTD'] / data2['Units_MTD']);
  //           if (key === 'TotalGross_PVR') return (data2['TotalGross_MTD'] / data2['Units_MTD']);

  //           if (key === 'Units_Diff' && data2['Units_Pace'] != 0 && data2['Units_Pace'] != null && data2['Units_Target'] != 0 && data2['Units_Target'] != null) return (data2['Units_Pace'] - data2['Units_Target']);
  //           if (key === 'FrontGross_Diff' && data2['FrontGross_Pace'] != 0 && data2['FrontGross_Pace'] != null && data2['FrontGross_Target'] != 0 && data2['FrontGross_Target'] != null) return (data2['FrontGross_Pace'] - data2['FrontGross_Target']);
  //           if (key === 'BackGross_Diff' && data2['BackGross_Pace'] != 0 && data2['BackGross_Pace'] != null && data2['BackGross_Target'] != 0 && data2['BackGross_Target'] != null) return (data2['BackGross_Pace'] - data2['BackGross_Target']);
  //           if (key === 'TotalGross_Diff' && data2['TotalGross_Pace'] != 0 && data2['TotalGross_Pace'] != null && data2['TotalGross_Target'] != 0 && data2['TotalGross_Target'] != null) return (data2['TotalGross_Pace'] - data2['TotalGross_Target']);

  //           const val = data2[key];
  //           return val === 0 || val == null ? '-' : val;
  //         });
  //         const nestedRow = worksheet.addRow(nestedRowData);

  //         bindingHeaders.forEach((key, index) => {
  //           const cell = nestedRow.getCell(index + 1);
  //           if (currencyFields.includes(key) && typeof cell.value === 'number') {
  //             cell.numFmt = '"$"#,##0';
  //             cell.alignment = { horizontal: 'right' };
  //             if (cell.value < 0) {
  //               cell.font = { color: { argb: 'FFFF0000' }, }
  //             }
  //           } else if (!isNaN(Number(cell.value))) {
  //             cell.alignment = { horizontal: 'right' };
  //           }
  //         });


  //         if (data2.SubData != undefined) {
  //           for (const data3 of data2.SubData) {
  //             const nestedRowData = bindingHeaders.map(key => {
  //               if (key === 'data1') return '          ' + capitalize(data3['data3']);
  //               if (key === 'FrontGross_PVR') return (data3['FrontGross_MTD'] / data3['Units_MTD']);
  //               if (key === 'BackGross_PVR') return (data3['BackGross_MTD'] / data3['Units_MTD']);
  //               if (key === 'TotalGross_PVR') return (data3['TotalGross_MTD'] / data3['Units_MTD']);

  //               if (key === 'Units_Target') return '';
  //               if (key === 'FrontGross_Target') return '';
  //               if (key === 'BackGross_Target') return '';
  //               if (key === 'TotalGross_Target') return '';

  //               if (key === 'FrontGross_Diff') return '';
  //               if (key === 'BackGross_Diff') return '';
  //               if (key === 'TotalGross_Diff') return '';

  //               const val = data3[key];
  //               return val === 0 || val == null ? '-' : val;
  //             });
  //             const nestedRow = worksheet.addRow(nestedRowData);

  //             bindingHeaders.forEach((key, index) => {
  //               const cell = nestedRow.getCell(index + 1);
  //               if (currencyFields.includes(key) && typeof cell.value === 'number') {
  //                 cell.numFmt = '"$"#,##0';
  //                 cell.alignment = { horizontal: 'right' };
  //                 if (cell.value < 0) {
  //                   cell.font = { color: { argb: 'FFFF0000' }, }
  //                 }
  //               } else if (!isNaN(Number(cell.value))) {
  //                 cell.alignment = { horizontal: 'right' };
  //               }
  //             });
  //           }
  //         }
  //       }
  //     }
  //   }
  //   worksheet.columns.forEach((column: any) => {
  //     let maxLength = 20;
  //     column.width = maxLength + 2;
  //   });
  //   workbook.xlsx.writeBuffer().then((data: any) => {
  //     const blob = new Blob([data], {
  //       type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  //     });
  //     this.shared.exportToExcel(workbook, 'Sales Gross')

  //   });
  // }

  exportToExcelBackGross() {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Back Gross');

    const from = new Date(this.FromDate);
    const to = new Date(this.ToDate);

    const monthText =
      from.toLocaleString('default', { month: 'long' }) +
      (from.getMonth() !== to.getMonth()
        ? ' - ' + to.toLocaleString('default', { month: 'long' })
        : '');

    const dateText =
      from.getDate() +
      (from.getFullYear() !== to.getFullYear() ? ' ' + from.getFullYear() : '') +
      ' - ' +
      to.getDate() +
      (from.getFullYear() !== to.getFullYear()
        ? ' ' + to.getFullYear()
        : ', ' + from.getFullYear());

    let dateHeader = '';
    if (this.datetype() === 'C') {
      const fromText =
        ('0' + (from.getMonth() + 1)).slice(-2) + '.' +
        ('0' + from.getDate()).slice(-2) + '.' +
        from.getFullYear();

      const toText =
        ('0' + (to.getMonth() + 1)).slice(-2) + '.' +
        ('0' + to.getDate()).slice(-2) + '.' +
        to.getFullYear();

      dateHeader = `${fromText}\n-\n${toText}`;
    } else {
      dateHeader = this.datetype();
    }

    /* ================= HEADER ================= */
    worksheet.addRow([
      monthText,
      'Units',
      'Back Gross', '', '', '', '',
      'Finance Sales', '', '', '',
      'Product Sales', '', '', '', ''
    ]);

    worksheet.mergeCells('B1:B1');
    worksheet.mergeCells('C1:G1');
    worksheet.mergeCells('H1:K1');
    worksheet.mergeCells('L1:P1');

    worksheet.addRow([
      dateText,
      dateHeader,
      dateHeader, 'Pace', 'Target', '+/-', 'PVR',
      'Gross', 'PVR', 'Count', 'Pen %',
      'Gross', 'PVR', 'Count', 'Pen %', 'Per Trans'
    ]);

    /* ================= HEADER STYLE ================= */
    [1, 2].forEach(r => {
      worksheet.getRow(r).eachCell(cell => {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: r === 1 ? '0554EF' : '4584FF' }
        };
        cell.font = { bold: true, color: { argb: 'FFFFFF' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.border = {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' }
        };
      });
    });

    /* ================= FORMAT FUNCTION ================= */
    const formatRow = (row: any, level = 0) => {
      row.eachCell((cell: any, colNumber: number) => {

        if (colNumber === 1) {
          cell.alignment = {
            horizontal: 'left',
            vertical: 'middle',
            indent: level * 2
          };
          return;
        }

        if (typeof cell.value === 'number') {

          // Units + Count
          if (colNumber === 2 || colNumber === 10 || colNumber === 14) {
            cell.numFmt = '#,##0';
          }

          // Pen %
          else if (colNumber === 11 || colNumber === 15) {
            cell.numFmt = '0%';
            cell.value = cell.value / 100;
          }

          // Per Trans
          else if (colNumber === 16) {
            cell.numFmt = '0.00';
          }

          // Currency
          else {
            cell.numFmt = '"$" * #,##0;[Red]"$" * -#,##0';
          }

          if (cell.value < 0) {
            cell.font = { color: { argb: 'FF0000' } };
          }
        }

        cell.alignment = {
          horizontal: 'right',
          vertical: 'middle'
        };
      });
    };

    /* ================= DATA ================= */
    this.BackGross.forEach((lvl1: any) => {

      const row1 = worksheet.addRow([
        lvl1.data1 || '-',
        lvl1.Units_MTD || 0,

        lvl1.BackGross_MTD || 0,
        lvl1.BackGross_Pace || 0,
        lvl1.BackGross_Target || 0,
        lvl1.BackGross_Diff || 0,
        lvl1.BackGross_PVR || 0,

        lvl1.figross || 0,
        lvl1.FigrossPVR || 0,
        lvl1.FRCOUNT || 0,
        lvl1.FIPen || 0,

        lvl1.ProductSale || 0,
        lvl1.ProductPVR || 0,
        lvl1.productdealcount || 0,
        lvl1.Productpen || 0,
        lvl1.peorductPertra || 0
      ]);

      row1.outlineLevel = 0;

      // ✅ Level 0 color
      row1.eachCell(cell => {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'D9E7FF' }
        };
      });

      // ✅ REPORT TOTAL
      if (lvl1.data1 === 'REPORTS TOTAL') {
        row1.eachCell(cell => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: '8DB4FF' }
          };
          cell.font = { bold: true };
        });
      }

      formatRow(row1, 0);

      /* LEVEL 2 */
      if (Array.isArray(lvl1.data2)) {
        lvl1.data2.forEach((lvl2: any) => {

          const row2 = worksheet.addRow([
            lvl2.data2 || '-',
            lvl2.Units_MTD || 0,

            lvl2.BackGross_MTD || 0,
            lvl2.BackGross_Pace || 0,
            lvl2.BackGross_Target || 0,
            lvl2.BackGross_dif || 0,
            lvl2.BackGross_PVR || 0,

            lvl2.figross || 0,
            lvl2.FigrossPVR || 0,
            lvl2.FRCOUNT || 0,
            lvl2.FIPen || 0,

            lvl2.ProductSale || 0,
            lvl2.ProductPVR || 0,
            lvl2.productdealcount || 0,
            lvl2.Productpen || 0,
            lvl2.peorductPertra || 0
          ]);

          row2.outlineLevel = 1;
          formatRow(row2, 1);

          /* LEVEL 3 */
          if (Array.isArray(lvl2.SubData)) {
            lvl2.SubData.forEach((lvl3: any) => {

              const row3 = worksheet.addRow([
                lvl3.data3 || '-',
                lvl3.Units_MTD || 0,

                lvl3.BackGross_MTD || 0,
                lvl3.BackGross_Pace || 0,
                '', '', '',

                lvl3.figross || 0,
                lvl3.FigrossPVR || 0,
                lvl3.FRCOUNT || 0,
                lvl3.FIPen || 0,

                lvl3.ProductSale || 0,
                lvl3.ProductPVR || 0,
                lvl3.productdealcount || 0,
                lvl3.Productpen || 0,
                lvl3.peorductPertra || 0
              ]);

              row3.outlineLevel = 2;
              formatRow(row3, 2);
            });
          }

        });
      }

    });

    /* ================= BORDERS ================= */
    worksheet.eachRow(row => {
      row.eachCell(cell => {
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      });
    });

    /* ================= FREEZE ================= */
    worksheet.views = [{ state: 'frozen', xSplit: 1, ySplit: 2 }];

    /* ================= WIDTH ================= */
    worksheet.columns.forEach((col, i) => {
      col.width = i === 0 ? 35 : 15;
    });

    worksheet.properties.outlineLevelRow = 2;

    /* ================= DOWNLOAD ================= */
    workbook.xlsx.writeBuffer().then(data => {
      saveAs(new Blob([data]), 'BackGross.xlsx');
    });
  }


  exportToExcelSalesGross() {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Sales Gross');
    /* ================= HEADER ROW 1 ================= */
    const from = new Date(this.FromDate);
    const to = new Date(this.ToDate);
    const monthText =
      from.toLocaleString('default', { month: 'long' }) +
      (from.getMonth() !== to.getMonth()
        ? ' - ' + to.toLocaleString('default', { month: 'long' })
        : '');

    const dateText =
      from.getDate() +
      (from.getFullYear() !== to.getFullYear() ? ' ' + from.getFullYear() : '') +
      ' - ' +
      to.getDate() +
      (from.getFullYear() !== to.getFullYear()
        ? ' ' + to.getFullYear()
        : ', ' + from.getFullYear());
    let dateHeader = '';

    if (this.datetype() === 'C') {
      const fromText =
        ('0' + (from.getMonth() + 1)).slice(-2) + '.' +
        ('0' + from.getDate()).slice(-2) + '.' +
        from.getFullYear();

      const toText =
        ('0' + (to.getMonth() + 1)).slice(-2) + '.' +
        ('0' + to.getDate()).slice(-2) + '.' +
        to.getFullYear();

      dateHeader = `${fromText}\n-\n${toText}`; // same as HTML (with line break)
    } else {
      dateHeader = this.datetype();
    }
    worksheet.addRow([
      monthText,
      'Units', '', '', '', '',
      'Front Gross', '', '', '', '',
      'Back Gross', '', '', '', '',
      'Total Gross', '', '', '', ''
    ]);

    // worksheet.mergeCells('A1:A2');
    worksheet.mergeCells('B1:F1');
    worksheet.mergeCells('G1:K1');
    worksheet.mergeCells('L1:P1');
    worksheet.mergeCells('Q1:U1');

    /* ================= HEADER ROW 2 ================= */
    worksheet.addRow([
      dateText,
      dateHeader, 'Pace', 'Target', '+/-', 'Per Day',
      dateHeader, 'Pace', 'Target', '+/-', 'PVR',
      dateHeader, 'Pace', 'Target', '+/-', 'PVR',
      dateHeader, 'Pace', 'Target', '+/-', 'PVR'
    ]);

    /* ================= HEADER STYLE ================= */
    [1, 2].forEach(r => {
      worksheet.getRow(r).eachCell(cell => {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: r === 1 ? '0554EF' : '4584FF' }
        };
        cell.font = { bold: true, color: { argb: 'FFFFFF' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.border = {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' }
        };
      });
    });

    /* ================= FORMAT FUNCTION ================= */
    const formatRow = (row: any, level: number = 0) => {
      row.eachCell((cell: any, colNumber: number) => {

        if (colNumber === 1) {
          cell.alignment = {
            horizontal: 'left',
            vertical: 'middle',
            indent: level * 2
          };
          return;
        }

        if (typeof cell.value === 'number') {

          if (colNumber === 6) {
            cell.numFmt = '0.00';
          }
          else if (colNumber >= 7) {
            cell.numFmt = '"$" * #,##0;[Red]"$" * -#,##0';
          }
          else {
            cell.numFmt = '#,##0';
          }

          if (cell.value < 0) {
            cell.font = { color: { argb: 'FF0000' } };
          }
        }

        cell.alignment = {
          horizontal: 'right',
          vertical: 'middle'
        };
      });
    };

    /* ================= DATA ================= */
    this.SalesData.forEach((lvl1: any, i: number) => {

      const row1 = worksheet.addRow([
        lvl1.data1,
        lvl1.Units_MTD, lvl1.Units_Pace, lvl1.Units_Target, lvl1.Units_Diff, lvl1.PerDay,
        lvl1.FrontGross_MTD, lvl1.FrontGross_Pace, lvl1.FrontGross_Target, lvl1.FrontGross_Diff, lvl1.FrontGross_PVR,
        lvl1.BackGross_MTD, lvl1.BackGross_Pace, lvl1.BackGross_Target, lvl1.BackGross_Diff, lvl1.BackGross_PVR,
        lvl1.TotalGross_MTD, lvl1.TotalGross_Pace, lvl1.TotalGross_Target, lvl1.TotalGross_Diff, lvl1.TotalGross_PVR
      ]);

      row1.outlineLevel = 0;
      row1.eachCell(cell => {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'D9E7FF' }
        };
      });
      /* ===== REPORT TOTAL ===== */
      if (lvl1.data1 === 'REPORTS TOTAL') {
        row1.eachCell(cell => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: '8DB4FF' }
          };
          cell.font = { bold: true };
        });
      }

      // /* ===== Alternate Row ===== */
      // else if (i % 2 === 0) {
      //   row1.eachCell(cell => {
      //     cell.fill = {
      //       type: 'pattern',
      //       pattern: 'solid',
      //       fgColor: { argb: 'F9FBFF' }
      //     };
      //   });
      // }

      formatRow(row1, 0);

      /* ================= LEVEL 2 ================= */
      lvl1.Data2?.forEach((lvl2: any) => {

        const row2 = worksheet.addRow([
          lvl2.data2,
          lvl2.Units_MTD, lvl2.Units_Pace, lvl2.Units_Target,
          lvl2.Units_Pace - lvl2.Units_Target, lvl2.PerDay,
          lvl2.FrontGross_MTD, lvl2.FrontGross_Pace, lvl2.FrontGross_Target,
          lvl2.FrontGross_Pace - lvl2.FrontGross_Target, lvl2.FrontGross_MTD / (lvl2.Units_MTD || 1),
          lvl2.BackGross_MTD, lvl2.BackGross_Pace, lvl2.BackGross_Target,
          lvl2.BackGross_Pace - lvl2.BackGross_Target, lvl2.BackGross_MTD / (lvl2.Units_MTD || 1),
          lvl2.TotalGross_MTD, lvl2.TotalGross_Pace, lvl2.TotalGross_Target,
          lvl2.TotalGross_Pace - lvl2.TotalGross_Target, lvl2.TotalGross_MTD / (lvl2.Units_MTD || 1)
        ]);

        row2.outlineLevel = 1;

        formatRow(row2, 1);

        /* ================= LEVEL 3 ================= */
        lvl2.SubData?.forEach((lvl3: any) => {

          const row3 = worksheet.addRow([
            lvl3.data3,
            lvl3.Units_MTD, lvl3.Units_Pace, '', '',
            lvl3.PerDay,
            lvl3.FrontGross_MTD, lvl3.FrontGross_Pace, '', '',
            lvl3.FrontGross_MTD / (lvl3.Units_MTD || 1),
            lvl3.BackGross_MTD, lvl3.BackGross_Pace, '', '',
            lvl3.BackGross_MTD / (lvl3.Units_MTD || 1),
            lvl3.TotalGross_MTD, lvl3.TotalGross_Pace, '', '',
            lvl3.TotalGross_MTD / (lvl3.Units_MTD || 1)
          ]);

          row3.outlineLevel = 2;
          formatRow(row3, 2);
        });

      });

    });

    /* ================= BORDERS ================= */
    worksheet.eachRow(row => {
      row.eachCell(cell => {
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      });
    });

    /* ================= FREEZE ================= */
    worksheet.views = [
      { state: 'frozen', xSplit: 1, ySplit: 2 }
    ];

    /* ================= COLUMN WIDTH ================= */
    worksheet.columns.forEach((col, index) => {
      col.width = index === 0 ? 35 : 15;
    });

    /* ================= GROUPING ================= */
    worksheet.properties.outlineLevelRow = 2;

    /* ================= DOWNLOAD ================= */
    workbook.xlsx.writeBuffer().then(data => {
      saveAs(new Blob([data]), 'SalesGross.xlsx');
    });
  }


}
