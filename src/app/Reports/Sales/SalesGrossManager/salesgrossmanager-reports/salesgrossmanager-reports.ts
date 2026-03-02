import { Component, HostListener, Input, SimpleChanges } from '@angular/core';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { DateRangePicker } from '../../../../CommonFilters/date-range-picker/date-range-picker';
import { Stores } from '../../../../CommonFilters/stores/stores';
import { ToastContainer } from '../../../../Layout/toast-container/toast-container';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { Setdates } from '../../../../Core/Providers/SetDates/setdates';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';

@Component({
  selector: 'app-salesgrossmanager-reports',
  imports: [SharedModule, DateRangePicker, Stores,ToastContainer],
  templateUrl: './salesgrossmanager-reports.html',
  styleUrl: './salesgrossmanager-reports.scss',
})
export class SalesgrossmanagerReports {

  ngChanges: any = []

  // Report Popup
  @Input() header: any;
  groupId: any = [1];
  storeIds: any = [1]
  stores: any = []
  groupsArray: any = [];
  storename: any = ''
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  storesFilterData: any = {
    'groupsArray': this.groupsArray, 'groupId': this.groupId, 'storesArray': this.stores, 'storeids': this.storeIds, 'type': 'M', 'others': 'N',
    'groupName': this.groupName, 'storename': this.storename, storecount: null, 'storedisplayname': this.storedisplayname
  };
  Performance: string = 'Load';

  minDate!: Date;
  maxDate!: Date;
  DateType: any = 'MTD';
  displaytime: any = '';
  FromDate: any = '';
  ToDate: any = '';
  Dates: any = {
    'FromDate': this.FromDate, 'ToDate': this.ToDate, "MaxDate": this.maxDate, 'MinDate': this.minDate, 'DateType': this.DateType, 'DisplayTime': this.displaytime,
    Types: [
      { 'code': 'MTD', 'name': 'MTD' },
      { 'code': 'YTD', 'name': 'YTD' },
      { 'code': 'PYTD', 'name': 'PYTD' },
      { 'code': 'LM', 'name': 'Last Month' },
      { 'code': 'PM', 'name': 'Same Month PY' },
    ]
  }


  selectedDataGrouping: any = [];
  GroupingDetails: any = [];
  dataGrouping: any = [];
  datagrp: any = [];
  Groupingcols = [
    { id: 1, columnname: 'store' },
    { id: 2, columnname: 'ad_dealtype' },
    { id: 3, columnname: 'ad_dealtype2' },
    { id: 4, columnname: 'salesperson' },
    { id: 5, columnname: 'fimanager' },
    { id: 6, columnname: 'salesmanager' },
    { id: 16, columnname: 'ad_Lender' },
    { id: 17, columnname: 'ad_dealstatus' },
    { id: 19, columnname: 'ad_Year' },
    { id: 24, columnname: 'ad_styleid' },
    { id: 25, columnname: 'ad_Month' },

    { id: 7, columnname: '' },
    { id: 8, columnname: '' },
    { id: 9, columnname: '' },
    { id: 10, columnname: 'ad_custzip' },
    { id: 11, columnname: 'ad_custstate' },
    { id: 12, columnname: 'ad_custAge' },
    { id: 13, columnname: '' },
    { id: 14, columnname: '' },
    { id: 15, columnname: '' },
    { id: 18, columnname: '' },
    { id: 20, columnname: 'ad_make' },
    { id: 21, columnname: 'ad_model' },
    { id: 22, columnname: '' },
    { id: 23, columnname: 'ad_trim' },
  ];

  TeamsArray: any = [
    { name: 'Sales Persons', code: 'SP' },
    { name: 'Sales Managers', code: 'SM' },
    { name: 'F&I Managers', code: 'FM' },

  ]
  Teams = 'SP'
  salesPersons: any = [];
  salesManagers: any = [];
  financeManager: any = [];
  selectedSalespersonvalues: any = [];
  selectedSalesManagersvalues: any = [];
  selectedFiManagersvalues: any = [];

  overallSelectedpeople: any = 0
  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent) {
    const clickedInside = (event.target as HTMLElement).closest('.dropdown-toggle, .reportstores-card , .timeframe');
    if (!clickedInside) {
      this.activePopover = -1;
    }
  }

  ngOnChanges(changes: SimpleChanges) {
    console.log(changes, 'Report');
    this.ngChanges = changes['header'].currentValue[0];
    this.FromDate = this.ngChanges.fromDate;
    this.ToDate = this.ngChanges.toDate;
    this.DateType = this.ngChanges.datevaluetype;
    this.groupId = this.ngChanges.groups
    this.setDates(this.ngChanges.datevaluetype)
    // this.StoresData(this.ngChanges)
    this.neworused = []
    this.neworused = this.ngChanges.dealType;
    this.retailorlease = [];
    this.retailorlease = this.ngChanges.saleType
    this.dealstatus = [];
    this.dealstatus = this.ngChanges.dealStatus
    this.reporttotal = this.ngChanges.ReportTotal
    if (changes['header'].previousValue != undefined) {
      this.datagrp = []
      this.dataGrouping.forEach((ele: any) => {
        ele.state = false
        this.shared.api.GetHeaderData().subscribe((res) => {
          if (
            res.obj.title == 'Sales Gross (Manager)'
          ) {
            this.GroupingDetails = [];
            this.GroupingDetails = res.obj;
            this.selectedDataGrouping = [];
            if (this.GroupingDetails.path1id != '') {
              if (ele.ARG_ID == this.GroupingDetails.path1id) {
                this.datagrp[0] = ele;
                ele.state = true;
              }
            }
            if (this.GroupingDetails.path12d != '') {
              if (ele.ARG_ID == this.GroupingDetails.path2id) {
                this.datagrp[1] = ele;
                ele.state = true;
              }
            }
            if (this.GroupingDetails.path3id != '') {
              if (ele.ARG_ID == this.GroupingDetails.path3id) {
                this.datagrp[2] = ele;
                ele.state = true;
              }
            }
          }
        });
      });
      this.selectedDataGrouping = this.datagrp;
    }
  }
  constructor(private shared: Sharedservice, private datesSrvc: Setdates,private toast:ToastService) { }

  ngOnInit() {
    this.getDataGroupings();
    this.getGroups();
    this.getEmployees('SP', '2', '1', 'Bar');
    this.getEmployees('F', '2', '1', 'Bar');
    this.getEmployees('M', '2', '1', 'Bar');
  }

  getDataGroupings() {
    this.selectedDataGrouping = [];
    this.GroupingDetails = [];
    const obj = {
      pageid: 1,
    };

    this.shared.api.postmethod(this.shared.common.routeEndpoint + 'GetDataGroupingsbyPage', obj).subscribe(
      (res) => {
        if (res.status == 200) {
          this.dataGrouping = res.response;
          // .sort((a, b) =>
          //   a.ARG_LABEL > b.ARG_LABEL ? 1 : -1
          // );
          this.dataGrouping.forEach((ele: any) => {
            ele.state = false;
            this.Groupingcols.forEach((val: any) => {
              if (ele.ARG_ID == val.id) {
                ele.columnName = val.columnname;
              }
              if (
                ele.ARG_ID != '1' &&
                ele.ARG_ID != '2' &&
             
           
                ele.ARG_ID != '20' &&
                ele.ARG_ID != '21' &&
              
                ele.ARG_ID != '3' &&
                ele.ARG_ID != '4' &&
                ele.ARG_ID != '5' &&
                ele.ARG_ID != '6' &&
             
                ele.ARG_ID != '17' &&

             
                ele.ARG_ID != '25'
              ) {
                ele.Active = 'N';
              } else {
                ele.Active = 'Y';
              }
            });
            this.shared.api.GetHeaderData().subscribe((res) => {
              if (
                res.obj.title == 'Sales Gross (Manager)' &&
                this.Performance == 'Load'
              ) {
                // // console.log(res.obj);

                this.GroupingDetails = [];
                this.GroupingDetails = res.obj;

                this.selectedDataGrouping = [];

                if (this.GroupingDetails.path1id != '') {
                  if (ele.ARG_ID == this.GroupingDetails.path1id) {
                    this.datagrp[0] = ele;

                    ele.state = true;
                  }
                }
                if (this.GroupingDetails.path12d != '') {
                  if (ele.ARG_ID == this.GroupingDetails.path2id) {
                    this.datagrp[1] = ele;
                    ele.state = true;
                  }
                }
                if (this.GroupingDetails.path3id != '') {
                  if (ele.ARG_ID == this.GroupingDetails.path3id) {
                    this.datagrp[2] = ele;
                    ele.state = true;
                  }
                }
              }
            });
          });
          this.selectedDataGrouping = this.datagrp;
          // // console.log(this.selectedDataGrouping);
        } else {
          this.toast.show('Invalid Details.', 'danger', 'Error');
        }
      },
      (error) => {
        // console.log(error);
      }
    );
  }
  SelectedData(val: any) {
    if (val.state == false) {
      if (this.selectedDataGrouping.length >= 3) {
       
        this.toast.show('Select up to 3 Filters only to Group your data', 'warning', 'Warning');
      } else {
        val.state = true;
        this.selectedDataGrouping.push(val);
      }
    } else {
      val.state = false;
      this.selectedDataGrouping.splice(
        this.selectedDataGrouping.indexOf(val),
        1
      );
    }

  }

  teamsselection(e: any) {
    this.Teams = e;
    // // console.log(this.Teams,this.activePopover);
    // // console.log(this.salesPersons, this.salesManagers, this.financeManager);


  }

  employees(block: any, e: any) {
    
    if (block == 'SP') {
      const index = this.selectedSalespersonvalues.findIndex(
        (i: any) => i == e
      );
      if (index >= 0) {
        this.selectedSalespersonvalues.splice(index, 1);
      } else {
        this.selectedSalespersonvalues.push(e);
      }
    }
    if (block == 'SM') {
      const index = this.selectedSalesManagersvalues.findIndex(
        (i: any) => i == e
      );
      if (index >= 0) {
        this.selectedSalesManagersvalues.splice(index, 1);
      } else {
        this.selectedSalesManagersvalues.push(e);
      }
    }
    if (block == 'FM') {
      const index = this.selectedFiManagersvalues.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.selectedFiManagersvalues.splice(index, 1);
      } else {
        this.selectedFiManagersvalues.push(e);
      }
    }
    if (block == 'AllSP') {

      if (e == 1) {
        this.selectedSalespersonvalues = [];
      } else if (e == 0) {
        this.selectedSalespersonvalues = this.salesPersons.map(function (
          a: any
        ) {
          return a.SPID;
        });
      }
      // // console.log(this.selectedSalespersonvalues);
    }
    if (block == 'AllSM') {
      if (e == 1) {
        this.selectedSalesManagersvalues = [];
      } else if (e == 0) {
        this.selectedSalesManagersvalues = this.salesManagers.map(function (
          a: any
        ) {
          return a.SmId;
        });
      }
      // // console.log(this.selectedSalesManagersvalues);
    }
    if (block == 'AllFM') {
      if (e == 1) {
        this.selectedFiManagersvalues = [];
      } else if (e == 0) {
        this.selectedFiManagersvalues = this.financeManager.map(function (
          a: any
        ) {
          return a.FiId;
        });
      }
      // // console.log(this.selectedFiManagersvalues);
    }
    this.overallSelectedpeople = this.selectedFiManagersvalues.length + this.selectedSalesManagersvalues.length + this.selectedSalespersonvalues.length

  }
  spinnerLoaderteams: boolean = false
  getEmployees(val: any, ids: any, count: any, barorpopup?: any) {
    const obj = {
      AS_ID: this.storeIds,
      type: val,
    };
    this.spinnerLoaderteams = true
    this.shared.api.postmethod(this.shared.common.routeEndpoint + 'GetEmployeesDev', obj).subscribe(
      (res) => {
        if (res.status == 200) {
          // //// // console.log('fjdgdgdfgdfhkl');

          if (val == 'SP') {
            this.salesPersons = res.response.filter(
              (e: any) => e.SPName != 'Unknown' && e.SPName != ''
            );
            this.selectedSalespersonvalues = this.salesPersons.map(function (
              a: any
            ) {
              return a.SPID;
            });
            if (barorpopup == 'Bar') {
              if (this.ngChanges.sp != '') {
                let spids = this.ngChanges.sp.split(',');
                this.selectedSalespersonvalues = spids;
              }
              if (this.ngChanges.sp == '0') {
                this.selectedSalespersonvalues = this.salesPersons.map(function (
                  a: any
                ) {
                  return a.SPID;
                });
              }
              if (this.ngChanges.sp == '') {
                this.selectedSalespersonvalues = []
              }
            }

          }
          if (val == 'F') {
            this.financeManager = res.response.filter(
              (e: any) => e.FiName != 'Unknown'
            );
            this.selectedFiManagersvalues = this.financeManager.map(function (
              a: any
            ) {
              return a.FiId;
            });

            if (barorpopup == 'Bar') {
              if (this.ngChanges.fm != '') {
                let fiids = this.ngChanges.fm.split(',');
                this.selectedFiManagersvalues = fiids;
              }
              if (this.ngChanges.fm == '0') {
                this.selectedFiManagersvalues = this.financeManager.map(function (
                  a: any
                ) {
                  return a.FiId;
                });
              }
              if (this.ngChanges.fm == '') {
                this.selectedFiManagersvalues = []
              }
            }
          }
          if (val == 'M') {
            this.salesManagers = res.response.filter(
              (e: any) => e.SmName != 'Unknown'
            );
            this.selectedSalesManagersvalues = this.salesManagers.map(function (
              a: any
            ) {
              return a.SmId;
            });

            if (barorpopup == 'Bar') {
              if (this.ngChanges.sm != '') {
                let smids = this.ngChanges.sm.split(',');
                this.selectedSalesManagersvalues = smids;

              }
              if (this.ngChanges.sm == '0') {
                this.selectedSalesManagersvalues = this.salesManagers.map(function (
                  a: any
                ) {
                  return a.SmId;
                });
              }
              if (this.ngChanges.sm == '') {
                this.selectedSalesManagersvalues = []
              }
            }
            this.spinnerLoaderteams = false;
          }
          this.overallSelectedpeople = this.selectedFiManagersvalues.length + this.selectedSalesManagersvalues.length + this.selectedSalespersonvalues.length

        } else {
        
          this.toast.show('Invalid Details.', 'danger', 'Error');

        }
      },
      (error) => {
        // //// // console.log(error);
      }
    );
  }
  // Report popup Code
  activePopover: number = -1;
  togglePopover(popoverIndex: number) {
   
    if (this.activePopover === popoverIndex) {
      this.activePopover = -1;
    } else {
      this.activePopover = popoverIndex;
    }
  }
  ngAfterViewInit() {
    this.shared.api.getStores().subscribe((res: any) => {
      // console.log(this.shared.common.pageName);

      if (this.shared.common.pageName == 'Sales Gross (Manager)') {
        if (res.obj.storesData != undefined) {
          this.groupsArray = res.obj.storesData;
          this.groupId = this.ngChanges.groups;
          this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
          this.storeIds = this.ngChanges.storeIds;
          this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_Name : this.groupName = ''
          this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
          // console.log(this.stores, this.groupsArray, 'Stores and Groups');
          this.getStoresandGroupsValues()
          // // console.log(this.groupId,'....');
          // this.StoresData(this.ngChanges)

        }
      }
    })

  }
  scrollPosition = 0;

  getScrollPosition(event: any): void {
    this.scrollPosition = event.target.scrollLeft ;
    console.log(this.scrollPosition,event.target.scrollTop);
    
  }
  getGroups() {
    // console.log(this.shared.common.pageName, this.shared.common.groupsandstores);

    if (this.shared.common.groupsandstores != undefined) {
      if (this.shared.common.groupsandstores.length > 0) {
        this.groupsArray = this.shared.common.groupsandstores.filter((val: any) => val.sg_id != this.shared.common.reconID);
        this.groupId = this.ngChanges.groups
        this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
        this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_name : this.groupName = ''
        this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
        // console.log(this.stores, this.groupsArray, 'Stores and Groups');
        this.getStoresandGroupsValues()
        // this.StoresData(this.ngChanges)
      }
    }
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



    // this.storesFilterData = { ...this.storesFilterData, newProp: 'updated' }
    console.log(this.storesFilterData, 'Store FIlter Data');

    let allstrids = [];
    allstrids = [...this.storeIds]
    this.getEmployees('SP', allstrids.toString(), '1', 'Bar');
    this.getEmployees('F', allstrids.toString(), '1', 'Bar');
    this.getEmployees('M', allstrids.toString(), '1', 'Bar');
  }
  StoresData(data: any) {

    this.storeIds = data.storeids;
    this.groupId = data.groupId;
    this.storename = data.storename;
    this.groupName = data.groupName;
    this.storecount = data.storecount;
    this.storedisplayname = data.storedisplayname;
    // console.log(data, this.storeIds, this.groupId, this.storename, this.groupName, this.stores, this.groupsArray, 'Stores related data');

    let allstrids = [];
    allstrids = [...this.storeIds]
    this.getEmployees('SP', allstrids.toString(), '2', '');
    this.getEmployees('F', allstrids.toString(), '2', '');
    this.getEmployees('M', allstrids.toString(), '2', '');

  }
  setDates(type: any) {
    // localStorage.setItem('time', type);
    // this.datevaluetype=
    // console.log(type);
    if (type != 'C') {
      this.displaytime = '( ' + this.Dates.Types.filter((val: any) => val.code == type)[0].name + ' )';
    }
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
  updatedDates(data: any) {
    // console.log(data);
    this.FromDate = data.FromDate;
    this.ToDate = data.ToDate;
    this.DateType = data.DateType;
    this.displaytime = data.DisplayTime
  }
  neworused: any = ['New', 'Used'];
  retailorlease: any = ['Retail ', 'Lease ', 'Misc '];
  dealTypeData: any = ['All', 'Retail and Lease', 'Retail', 'Lease', 'Wholesale', 'Misc', 'Fleet', 'Demo', 'Special Order', 'Rental', 'Dealer Trade']

  includeorexclude: any = [''];
  dealstatus: any = ['Delivered', 'Capped', 'Finalized'];
  selecttarget: any = ['F'];
  reporttotal: any = 'T';
  includecharges: any = [''];
  Transactorgl: any = ['T'];
  multipleorsingle(block: any, e: any) {
    if (block == 'NU') {
      const index = this.neworused.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.neworused.splice(index, 1);
      } else {
        this.neworused.push(e);
      }
    }
    if (block == 'RL') {
      const index = this.retailorlease.findIndex((i: any) => i == e);
      if (index >= 0) {
        if (e == 'All') {
          // this.retailorlease.splice(index, 1);
          this.retailorlease = []
        } else {
          this.retailorlease.splice(index, 1);
          let allindex = this.retailorlease.findIndex((i: any) => i == 'All');
          if (allindex >= 0) {
            this.retailorlease.splice(allindex, 1);
          }
        }
      } else {
        if (e == 'All') {
          const dealdata = JSON.stringify(this.dealTypeData)
          this.retailorlease = JSON.parse(dealdata)
        } else {
          this.retailorlease.push(e);
          if (this.retailorlease.length == this.dealTypeData.length - 1) {
            const dealdata = JSON.stringify(this.dealTypeData)
            this.retailorlease = JSON.parse(dealdata)
          }
        }
      }
    }
    if (block == 'PH') {
      this.includeorexclude = [];
      this.includeorexclude.push(e);
    }
    if (block == 'DS') {
      const index = this.dealstatus.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.dealstatus.splice(index, 1);
      } else {
        this.dealstatus.push(e);
      }
    }
    if (block == 'ST') {
      const index = this.selecttarget.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.selecttarget.splice(index, 1);
      } else {
        this.selecttarget.push(e);
      }
    }
    if (block == 'RT') {
      this.reporttotal = e;
    }
    if (block == 'IC') {
      this.includecharges = [];
      this.includecharges.push(e);
    }
    if (block == 'SRC') {
      this.Transactorgl = [];
      this.Transactorgl.push(e);
    }
  }
  viewreport() {
    this.activePopover = -1
    let peoples = this.selectedFiManagersvalues.length + this.selectedSalesManagersvalues.length + this.selectedSalespersonvalues.length
    let aqusrc: any = ['All']
    if(!this.storeIds || this.storeIds.length === 0){
      this.toast.show(
        'Please Select Atleast One Store',
        'warning',
        'Warning'
      );
    }
    else if (this.selectedDataGrouping.length == 0) {
      this.toast.show(
        'Please Select Atleast One Grouping',
        'warning',
        'Warning'
      );
    } else if (this.dealstatus.length == 0) {
      this.toast.show(
        'Please Select Atleast One Deal Status',
        'warning',
        'Warning'
      );
     
    }
    else if (this.retailorlease.length == 0) {
      this.toast.show(
        'Please Select Atleast One Deal Type',
        'warning',
        'Warning'
      );
     
    }
    else if (peoples == 0) {
      this.toast.show(
        'Please Select Atleast One from People Filter',
        'warning',
        'Warning'
      );
      
    }
  
    else {
      const data = {
        Reference: 'Sales Gross (Manager)',
        FromDate: this.shared.datePipe.transform(this.FromDate, 'MM-dd-yyyy'),
        ToDate: this.shared.datePipe.transform(this.ToDate, 'MM-dd-yyyy'),
        // TotalReport: this.toporbottom[0],
        storeValues: this.storeIds.toString(),
        Spvalues:
          this.selectedSalespersonvalues.length == this.salesPersons.length
            ? '0'
            : this.selectedSalespersonvalues.toString(),
        SMvalues:
          this.selectedSalesManagersvalues.length == this.salesManagers.length
            ? '0'
            : this.selectedSalesManagersvalues.toString(),
        FIvalues:
          this.selectedFiManagersvalues.length == this.financeManager.length
            ? '0'
            : this.selectedFiManagersvalues.toString(),
        dealType: this.neworused,
        saleType: this.retailorlease,
        dealStatus: this.dealstatus,
        dataGroupingvalues: this.selectedDataGrouping,
        dateType: this.DateType,
        reportTotal: this.reporttotal,
        
        // gv: this.gridview,
        // acquisition: this.Acquisition.length == this.acquisitionsource.length ? aqusrc : this.Acquisition,
        groups: this.groupId,
        // otherstoreids: this.selectedotherstoreids
      };
      // // console.log(data);
      this.shared.api.SetReports({
        obj: data,
      });
      this.close();
      // }
      // }
    }
  }

  close() {
    this.Performance = 'Unload';
  }
}