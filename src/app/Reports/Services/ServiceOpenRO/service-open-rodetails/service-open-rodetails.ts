import { Component, OnInit, ViewChild, ElementRef, HostListener, Input } from '@angular/core';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { Setdates } from '../../../../Core/Providers/SetDates/setdates';
import { common } from '../../../../common';
import { Subscription } from 'rxjs';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
import { CurrencyPipe } from '@angular/common';
import { NgbActiveModal } from '@ng-bootstrap/ng-bootstrap';
import { Notes } from '../../../../Layout/notes/notes';

@Component({
  selector: 'app-service-open-rodetails',
  imports: [SharedModule,Notes],
  templateUrl: './service-open-rodetails.html',
  styleUrl: './service-open-rodetails.scss',
})
export class ServiceOpenRODetails {
  @Input() RODetailsObjectMain: any = [];
  NoData!: boolean;
  spinnerLoader: boolean = false;
  spinnerLoadersec: boolean = false;
  pageNumber: any = 0;
  ServiceData: any = [];
  callLoadingState: any = 'FL';
  QISearchName: any = '';
  constructor(
    public shared: Sharedservice, public setdates: Setdates, private comm: common, private cp: CurrencyPipe, private toast: ToastService,private ngbmodalActive:NgbActiveModal
  ) {

  }
  RODetailsObject: any = []
  ngOnInit(): void {
    this.RODetailsObject = [...this.RODetailsObjectMain]
    console.log(this.RODetailsObject[0]);
    if (this.RODetailsObject[0]) {
      this.total = this.RODetailsObject[0].Total
      this.totalPageCount = this.RODetailsObject[0].Total / 100;
      if (this.totalPageCount == Math.floor(this.totalPageCount)) {
      } else {
        this.totalPageCount = Math.floor(this.totalPageCount) + 1
      }
      if (this.RODetailsObject[0].topfive == 'Y') {
        this.ServiceData = this.RODetailsObject[0].data.RoInfo;
        this.details = this.RODetailsObject[0].data.RoInfo;
        this.details = this.details.map((v: any) => ({
          ...v, comment: '+', notesView: '+', completeDetails: this.RODetailsObject[0]
        }));
        this.details.some(function (x: any) {
          if (x.Notes != undefined && x.Notes != '') {
            x.Notes = JSON.parse(x.Notes);
          }
        });
        this.ServiceData = [...this.details]
        this.filterData();

      } else {
        this.spinnerLoader = true;

        this.getDetails();
      }
    }

  }
  details: any = [];
  viewRO(roData: any) {
    // const modalRef = this.shared.ngbmodal.open(RepairOrderComponent, { size: 'md', windowClass: 'compModal' });
    // modalRef.componentInstance.data = { ro: roData.ronumber, storeid: roData.StoreID, vin: roData.vin, vehicleid: roData.vehicleid, custno: roData.customernumber }; // Pass data to the modal component    
    // modalRef.result.then((result) => {
    //   console.log(result); // Handle modal close result
    // }, (reason) => {
    //   console.log(`Dismissed: ${reason}`); // Handle dismiss reason
    // });
  }
  total: any = 0
  totalPageCount: any = 0;
  clickedPage: number | null = null;
  currentPage: number = 1;
  itemsPerPage: number = 100;

  nextPage() {
    if (this.currentPage < this.getMaxPageNumber()) {
      this.currentPage++;
      this.clickedPage = null;
    }
    this.CurrentPageSetting = 1
  }

  prevPage() {
    if (this.currentPage > 1) {
      this.currentPage--;
      this.clickedPage = null;
    }
    this.CurrentPageSetting = 1
  }

  goToPage(page: number) {
    this.currentPage = page;
    this.clickedPage = null;
    this.CurrentPageSetting = 1;
  }

  goToFirstPage() {
    this.currentPage = 1;
    this.clickedPage = null;
    this.CurrentPageSetting = 1

  }
  goToLastPage() {
    this.currentPage = this.getMaxPageNumber();
    this.clickedPage = null;
    this.CurrentPageSetting = 1

  }
  getStartRecordIndex(): number {
    return (this.currentPage - 1) * this.itemsPerPage;
  }
  getEndRecordIndex(): number {
    const endIndex = this.getStartRecordIndex() + this.itemsPerPage;
    return endIndex > this.ServiceData.length ? this.ServiceData.length : endIndex;
  }
  getMaxPageNumber(): number {
    // console.log(Math.ceil(this.total / 100)-1);
    return Math.ceil(this.filteredServiceOpenROData.length / this.itemsPerPage);

    // return Math.ceil(this.total / 100);
  }
  getDetails() {
    this.spinnerLoader = true;
    const obj = {
      startdealdate: this.RODetailsObject[0].StartDate,
      enddealdate: this.RODetailsObject[0].EndDate,
      var1: this.RODetailsObject[0].var1,
      var2: this.RODetailsObject[0].var2,
      var3: '',
      var1Value: this.RODetailsObject[0].topfive == 'Y' ? this.selectedRO.completeDetails.data.data1 : this.RODetailsObject[0].var1Value,
      var2Value: this.RODetailsObject[0].topfive == 'Y' ? this.selectedRO.completeDetails.data.data2 : this.RODetailsObject[0].var2Value,
      var3Value: '',
      GrossTypeLabor:
        this.RODetailsObject[0].GrossTypeLabor,
      GrossTypeParts:
        this.RODetailsObject[0].GrossTypeParts,
      GrossTypeMisc: this.RODetailsObject[0].GrossTypeMisc,
      GrossTypeSublet:
        this.RODetailsObject[0].GrossTypeSublet,
      PaytypeCP: this.RODetailsObject[0].PaytypeCP,
      PaytypeWarranty: this.RODetailsObject[0].PaytypeWarranty,
      PaytypeInternal: this.RODetailsObject[0].PaytypeInternal,
      PaytypeExtendedWarranty: this.RODetailsObject[0].PaytypeExtendedWarranty,
      inventory: this.RODetailsObject[0].inventory,
      PageNumber: 0,
      PageSize: this.total,
      minage: this.RODetailsObject[0].AgeFrom,
      maxage: this.RODetailsObject[0].AgeTo,
      ROSTATUS: this.RODetailsObject[0].ROSTATUS,
      Oldro: this.RODetailsObject[0].topfive
    }

    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetServicesGrossSummaryDetailsV1Open', obj).subscribe(
      (res) => {
        //console.log(res);
        if (res.status == 200) {

          this.ngbmodalActive.dismiss()
          if (this.popupOpen == 'Y') {
            this.toast.show('Comment added Successfully!', 'success','Success');
          }
          this.popupOpen = 'N'
          if (res.response != undefined) {
            this.details = res.response.map((v: any) => ({
              ...v, comment: '+', notesView: '+'
            }));
            this.details.some(function (x: any) {
              if (x.Notes == null) {
                x.Notes = []
              }
              if (x.Notes != undefined && x.Notes != '') {
                x.Notes = JSON.parse(x.Notes);
                // x.LatestComment = x.LatestComment.map((v: any) => ({
                //   ...v,
                //   comment: '+'

                // }));


              }

            });
            this.ServiceData = this.details
            this.filterData();

            // [
            //   ...this.ServiceData,
            //   ...this.details,
            // ];
            this.callLoadingState == 'ANS' ? this.sort(this.column, this.ServiceData, this.callLoadingState) : ''

            this.spinnerLoader = false;
            this.spinnerLoadersec = false;
            if (this.ServiceData.length > 0) {
              this.NoData = false;
            } else {
              this.NoData = true;
            }
          } else {
            this.spinnerLoader = false;
            this.spinnerLoadersec = false;
            this.NoData = true;
          }

        }
        else {
          this.spinnerLoader = false;
          this.spinnerLoadersec = false;
          this.NoData = true;
        }
        // this.ServiceData.some(function (x: any) {
        //   if (x.TechDetails != undefined) {
        //     x.TechDetails = JSON.parse(x.TechDetails);
        //     // x.Data2 = x.Data2.map((v: any) => ({
        //     //   ...v,
        //     //   SubData: [],
        //     //   data2sign: '-',
        //     // }));
        //   }

        //   x.Dealer = '+'

        // if (path2 == '') {
        //   x.Dealer = '+';
        // } else {
        //   x.Dealer = '-';
        // }
        // });
        //console.log(this.ServiceData);

      })
  }
  get paginatedItems() {
    this.CurrentPageSetting != 1 ? this.currentPage = this.CurrentPageSetting : '';
    const startIndex = (this.currentPage - 1) * this.itemsPerPage;
    const endIndex = startIndex + this.itemsPerPage;
    return this.filteredServiceOpenROData.slice(startIndex, endIndex);
  }
  filteredServiceOpenROData: any = []
  filterData() {
    if (this.QISearchName.trim() !== '') {
   
      const search = this.QISearchName.toLowerCase();
      this.filteredServiceOpenROData = this.ServiceData
        .map((item: any) => {
          const matchesRonumber = item.ronumber?.toLowerCase().includes(search);
          const matchesOpenDate = item.opendate?.toLowerCase().includes(search);
          const matchesAge = typeof item.age === 'string' && item.age.toLowerCase().includes(search);
          const matchesRoStatus = item.RoStatus?.toLowerCase().includes(search);
          const matchesCName = item.CName?.toLowerCase().includes(search);
          const matchesVehicle = item.vehicle?.toLowerCase().includes(search);
          const matchesVinNo = item.vinno?.toLowerCase().includes(search);
          const matchesAdvisor = item.ServiceAdvisor_Name?.toLowerCase().includes(search);

          const matchesOtherFields =
            matchesRonumber || matchesOpenDate || matchesAge ||
            matchesRoStatus || matchesCName || matchesVehicle ||
            matchesVinNo || matchesAdvisor;

          if (matchesOtherFields) {
            return item; // Include as-is (keep all Notes)
          }

          // No primary field matched; check Notes
          if (Array.isArray(item.Notes)) {
            const filteredNotes = item.Notes.filter(
              (note: any) => note?.GN_Text?.toLowerCase().includes(search)
            );

            if (filteredNotes.length > 0) {
              return {
                ...item,
                Notes: filteredNotes // only matched notes
              };
            }
          }

          // Neither fields nor Notes matched — exclude
          return null;
        })
        .filter((item: any) => item !== null);
    } else {
      this.filteredServiceOpenROData = [...this.ServiceData];
    }
    this.callLoadingState == 'ANS' ? this.sort(this.column, this.filteredServiceOpenROData, this.callLoadingState) : ''
    let position = this.scrollpositionstoring + 10
    setTimeout(() => {
      this.scrollcent.nativeElement.scrollTop = position
      // //console.log(position);

    }, 500);
    this.pageNumber = 1;
  }
  toggleView(data: any) {
    if (data.notesView == '+') {
      data.notesView = '-'
    } else {
      data.notesView = '+'
    }
  }
  @ViewChild('scrollcent') scrollcent!: ElementRef;
  @ViewChild('scrollOne') scrollOne!: ElementRef;
  updateVerticalScroll(event: any): void {
   

  }

  popupOpen: any = ''
  notesViewState: boolean = true
  notesView() {
    this.notesViewState = !this.notesViewState
  }

  notesData: any = {}
  Notespopup: any;
  selectedRO: any = [];
  scrollpositionstoring: any = 0;
  scrollCurrentposition: any = 0;
  CurrentPageSetting: any = 1;

  addNotes(data: any, ref: any) {
    this.scrollpositionstoring = this.scrollCurrentposition

    this.selectedRO = data
    this.notesData = {
      store: data.StoreID,
      title1: data.ronumber,
      title2: '',
      apiRoute: 'AddGeneralNotes'
    }
    console.log(this.notesData);

    this.Notespopup = this.shared.ngbmodal.open(ref, { size: 'xxl', backdrop: 'static' });
  }

  closeNotes(e: any) {
    console.log(this.selectedRO);
    if (e == 'S') {
      this.callLoadingState = 'ANS'
      this.shared.ngbmodal.dismissAll();
      this.CurrentPageSetting = this.currentPage
      // this.selectedRO.Notes
      // if (this.RODetailsObject[0].topfive == 'Y') {
      //   // this.selectedRO.data.Notes= 
      //   this.ServiceData = this.RODetailsObject[0].data.RoInfo;
      //   this.details = this.RODetailsObject[0].data.RoInfo;
      //   this.filterData();
      // } else {
      // this.getDetails();
      // }
    }
    if (e == 'C') {
      this.shared.ngbmodal.dismissAll()
    }
  }

  savedNotes(e: any) {
    let obj = { "GN_Text": e.notes }
    this.selectedRO.Notes.unshift(obj)
    this.selectedRO.NotesStatus = 'Y';
    console.log(this.selectedRO);

  }

  isDesc: boolean = false;
  column: string = 'CategoryName';

  sort(property: any, data: any, state?: any) {
    if (state == undefined) {
      this.isDesc = !this.isDesc;
    }
    this.callLoadingState = 'FL'
    this.column = property;
    let direction = this.isDesc ? 1 : -1;
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

  exportAsXLSX() {
    this.shared.spinner.show()
    const obj = {
      startdealdate: this.RODetailsObject[0].StartDate,
      enddealdate: this.RODetailsObject[0].EndDate,
      var1: this.RODetailsObject[0].var1,
      var2: this.RODetailsObject[0].var2,
      var3: '',
      var1Value: this.RODetailsObject[0].topfive == 'Y' ? this.selectedRO.completeDetails.data.data1 : this.RODetailsObject[0].var1Value,
      var2Value: this.RODetailsObject[0].topfive == 'Y' ? this.selectedRO.completeDetails.data.data2 : this.RODetailsObject[0].var2Value,
      var3Value: '',
      GrossTypeLabor:
        this.RODetailsObject[0].GrossTypeLabor,
      GrossTypeParts:
        this.RODetailsObject[0].GrossTypeParts,
      GrossTypeMisc: this.RODetailsObject[0].GrossTypeMisc,
      GrossTypeSublet:
        this.RODetailsObject[0].GrossTypeSublet,
      PaytypeCP: this.RODetailsObject[0].PaytypeCP,
      PaytypeWarranty: this.RODetailsObject[0].PaytypeWarranty,
      PaytypeInternal: this.RODetailsObject[0].PaytypeInternal,
      PaytypeExtendedWarranty: this.RODetailsObject[0].PaytypeExtendedWarranty,

      inventory: this.RODetailsObject[0].inventory,
      PageNumber: 0,
      PageSize: '10000',
      minage: this.RODetailsObject[0].AgeFrom,
      maxage: this.RODetailsObject[0].AgeTo,
      ROSTATUS: this.RODetailsObject[0].ROSTATUS,
      Oldro: this.RODetailsObject[0].topfive

    }

    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetServicesGrossSummaryDetailsV1Open', obj).subscribe(
      (res) => {
        this.shared.spinner.hide()
        if (res.status == 200) {
          console.log(res);

          if (res.response != undefined) {
            if (res.response.length > 0) {
              let localarray = res.response;

              localarray.some(function (x: any) {
                if (x.Notes != undefined && x.Notes != '') {
                  x.Notes = JSON.parse(x.Notes);
                }

              });
              const workbook = this.shared.getWorkbook();
              const worksheet = workbook.addWorksheet('Service Open RO Details');
              worksheet.views = [
                {
                  state: 'frozen',
                  ySplit: 9, // Number of rows to freeze (2 means the first two rows are frozen)
                  topLeftCell: 'A10', // Specify the cell to start freezing from (in this case, the third row)
                  showGridLines: false,
                },
              ];
              const DateToday = this.shared.datePipe.transform(new Date(), 'MM/dd/yyyy h:mm:ss a');

              const titleRow = worksheet.getCell("A2"); titleRow.value = 'Service Open RO Details';
              titleRow.font = { name: 'Arial', family: 4, size: 15, bold: true };
              titleRow.alignment = { indent: 1, vertical: 'middle', horizontal: 'left' }



              const DateBlock = worksheet.getCell("L2"); DateBlock.value = DateToday;
              DateBlock.font = { name: 'Arial', family: 4, size: 10 };
              DateBlock.alignment = { vertical: 'middle', horizontal: 'center' }
              worksheet.addRow([''])
              const Store_Name = worksheet.addRow(['Store Name :']);
              Store_Name.getCell(1).font = { name: 'Arial', family: 4, size: 9, bold: true, };
              Store_Name.alignment = { indent: 1, vertical: 'middle', horizontal: 'left' }
              const StoreName = worksheet.getCell("B4"); StoreName.value = this.RODetailsObject[0].var1Value;
              StoreName.font = { name: 'Arial', family: 4, size: 9 };
              StoreName.alignment = { indent: 1, vertical: 'middle', horizontal: 'left' }

              const DATE_EXTENSION = this.shared.datePipe.transform(
                new Date(),
                'MMddyyyy'
              );

              const StartDealDate = worksheet.addRow(['Start Date :']);
              const startdealdate = worksheet.getCell('B5');
              startdealdate.value = this.RODetailsObject[0].StartDate;
              startdealdate.font = { name: 'Arial', family: 4, size: 9 };
              startdealdate.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
              StartDealDate.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
              StartDealDate.getCell(1).font = {
                name: 'Arial',
                family: 4,
                size: 9,
                bold: true,
              };
              const EndDealDate = worksheet.addRow(['End Date :']);
              EndDealDate.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
              const enddealdate = worksheet.getCell('B6');
              enddealdate.value = this.RODetailsObject[0].EndDate;
              enddealdate.font = { name: 'Arial', family: 4, size: 9 };
              enddealdate.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
              EndDealDate.getCell(1).font = {
                name: 'Arial',
                family: 4,
                size: 9,
                bold: true,
              };
              console.log(this.RODetailsObject[0].EndDate, this.RODetailsObject[0].StartDate, enddealdate.value, startdealdate.value);


              // const Var1Value = worksheet.addRow(['Advisor Name :']);
              // Var1Value.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
              // const var1value = worksheet.getCell('B7');
              // var1value.value = this.RODetailsObject[0].userName == '' ? '-' : this.RODetailsObject[0].userName == null ? '-' : this.RODetailsObject[0].userName;
              // var1value.font = { name: 'Arial', family: 4, size: 9 };
              // var1value.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
              // Var1Value.getCell(1).font = {
              //   name: 'Arial',
              //   family: 4,
              //   size: 9,
              //   bold: true,
              // };
              worksheet.addRow('');
              let Headings = [

                'RO #',
                'Date',
                'Age',
                'RO Status',
                'Customer',
                'Advisor',
                'Vehicle',
                'Stock',
                'Vin',
                'Total Gross',
                'Total Cost',
                'Total Sale',
                'Total GP%',
                'Total ELR',
                'Total Hours',
                'Total Discount',
                'Customer Pay	Gross',
                'Customer Pay	GP%',
                'Customer Pay	ELR',
                'Customer Pay	Hours',
                'Customer Pay	Discount',
                'Warranty Gross',
                'Warranty GP%',
                'Warranty ELR',
                'Warranty Hours',
                'Warranty Discount',

                // 'Extended Warranty Gross',
                // 'Extended Warranty GP%',
                // 'Extended Warranty ELR',
                // 'Extended Warranty Hours',
                // 'Extended Warranty Discount',

                'Internal Gross',
                'Internal GP%',
                'Internal ELR',
                'Internal Hours',
                'Internal Discount',
              ];
              const headerRow = worksheet.addRow(Headings);
              headerRow.font = {
                name: 'Arial',
                family: 4,
                size: 9,
                bold: true,
                color: { argb: 'FFFFFF' },
              };
              headerRow.height = 20;
              headerRow.alignment = {
                indent: 1,
                vertical: 'middle',
                horizontal: 'center',
              };
              headerRow.eachCell((cell, number) => {
                cell.fill = {
                  type: 'pattern',
                  pattern: 'solid',
                  fgColor: { argb: '2a91f0' },
                  bgColor: { argb: 'FF0000FF' },
                };
                headerRow.height = 20;
                cell.border = { right: { style: 'thin' } };
                cell.alignment = {
                  vertical: 'middle',
                  horizontal: 'center',
                  wrapText: true,
                };
              });
              var count = 0
              var notesCount = 9
              for (const d of localarray) {

                count++
                notesCount++
                d.opendate = this.shared.datePipe.transform(d.opendate, 'MM/dd/yyyy');
                const Data1 = worksheet.addRow([

                  d.ronumber == '' ? '-' : d.ronumber == null ? '-' : d.ronumber,
                  d.opendate == '' ? '-' : d.opendate == null ? '-' : d.opendate,
                  d.age == '' ? '-' : d.age == null ? '-' : d.age.toString(),
                  d.RoStatus == '' ? '-' : d.RoStatus == null ? '-' : d.RoStatus,
                  d.CName == '' ? '-' : d.CName == null ? '-' : d.CName,
                  d.ServiceAdvisor_Name == '' ? '-' : d.ServiceAdvisor_Name == null ? '-' : d.ServiceAdvisor_Name,
                  d.vehicle == '' ? '-' : d.vehicle == null ? '-' : d.vehicle,
                  d.stock == '' ? '-' : d.stock == null ? '-' : d.stock,
                  d.vinno == '' ? '-' : d.vinno == null ? '-' : d.vinno,
                  d.Totalgross == '' ? '-' : d.Totalgross == null ? '-' : parseFloat(d.Totalgross),
                  d.Totalcost == '' ? '-' : d.Totalcost == null ? '-' : parseFloat(d.Totalcost),
                  d.TotalSale == '' ? '-' : d.TotalSale == null ? '-' : parseFloat(d.TotalSale),
                  d.Retention == '' ? '-' : d.Retention == null ? '-' : parseFloat(d.Retention) + ' %',
                  d.TotalELR == '' ? '-' : d.TotalELR == null ? '-' : parseFloat(d.TotalELR),
                  d.Totalhours == '' ? '-' : d.Totalhours == null ? '-' : parseFloat(d.Totalhours),
                  d.Discount == '' ? '-' : d.Discount == null ? '-' : parseFloat(d.Discount),
                  d.CustomerPayGross == '' ? '-' : d.CustomerPayGross == null ? '-' : parseFloat(d.CustomerPayGross),
                  d.CustomerRetention == '' ? '-' : d.CustomerRetention == null ? '-' : parseFloat(d.CustomerRetention) + ' %',
                  d.CustomerPayELR == '' ? '-' : d.CustomerPayELR == null ? '-' : parseFloat(d.CustomerPayELR),
                  d.CustomerPayhours == '' ? '-' : d.CustomerPayhours == null ? '-' : parseFloat(d.CustomerPayhours),
                  d.DiscountCP == '' ? '-' : d.DiscountCP == null ? '-' : parseFloat(d.DiscountCP),
                  d.WarrantyGross == '' ? '-' : d.WarrantyGross == null ? '-' : parseFloat(d.WarrantyGross),
                  d.WarrantyRetention == '' ? '-' : d.WarrantyRetention == null ? '-' : parseFloat(d.WarrantyRetention) + ' %',
                  d.WarrantyELR == '' ? '-' : d.WarrantyELR == null ? '-' : parseFloat(d.WarrantyELR),
                  d.Warrantyhours == '' ? '-' : d.Warrantyhours == null ? '-' : parseFloat(d.Warrantyhours),
                  d.DiscountWP == '' ? '-' : d.DiscountWP == null ? '-' : parseFloat(d.DiscountWP),

                  // d.ExtendedGross == '' ? '-' : d.ExtendedGross == null ? '-' : parseFloat(d.ExtendedGross),
                  // d.ExtendedRetention == '' ? '-' : d.ExtendedRetention == null ? '-' : parseFloat(d.ExtendedRetention) + ' %',
                  // d.ExtendedELR == '' ? '-' : d.ExtendedELR == null ? '-' : parseFloat(d.ExtendedELR),
                  // d.Extenedhours == '' ? '-' : d.Extenedhours == null ? '-' : parseFloat(d.Extenedhours),
                  // d.DiscountEW == '' ? '-' : d.DiscountEW == null ? '-' : parseFloat(d.DiscountEW),


                  d.InternalGross == '' ? '-' : d.InternalGross == null ? '-' : parseFloat(d.InternalGross),
                  d.InternalRetention == '' ? '-' : d.InternalRetention == null ? '-' : parseFloat(d.InternalRetention) + ' %',
                  d.InternalELR == '' ? '-' : d.InternalELR == null ? '-' : parseFloat(d.InternalELR),
                  d.Internalhours == '' ? '-' : d.Internalhours == null ? '-' : parseFloat(d.Internalhours),
                  d.DiscountIP == '' ? '-' : d.DiscountIP == null ? '-' : parseFloat(d.DiscountIP),
                ]);
                // Data1.outlineLevel = 1; // Grouping level 1
                Data1.font = { name: 'Arial', family: 4, size: 8 };
                Data1.height = 18;
                // Data1.getCell(1).alignment = {indent: 1,vertical: 'middle', horizontal: 'left'}
                Data1.eachCell((cell, number) => {
                  cell.border = { right: { style: 'thin' } };
                  cell.numFmt = '$#,##0.00';
                  if (number == 13 || number == 18 || number == 23 || number == 28 || number == 33) {
                    cell.numFmt = '0.0';
                    cell.alignment = { indent: 1, vertical: 'middle', horizontal: 'center', };
                  }
                  if (number == 15 || number == 20 || number == 25 || number == 30 || number == 35) {
                    cell.numFmt = '0.00';
                    cell.alignment = { indent: 1, vertical: 'middle', horizontal: 'center', };
                  }




                  // else if (number == 15) {
                  //   cell.numFmt = '0.0';
                  //   cell.alignment = { indent: 1, vertical: 'middle', horizontal: 'center', };
                  // } else if (number == 20) {
                  //   cell.numFmt = '0.0';
                  //   cell.alignment = { indent: 1, vertical: 'middle', horizontal: 'center', };
                  // } else if (number == 25) {
                  //   cell.numFmt = '0.0';
                  //   cell.alignment = { indent: 1, vertical: 'middle', horizontal: 'center', };
                  // }
                  // if (number > 6 && number < 27) {
                  //   cell.alignment = { indent: 1, vertical: 'middle', horizontal: 'center', };
                  // } else if (number == 1) {
                  //   cell.numFmt = '###0';
                  //   cell.alignment = { indent: 1, vertical: 'middle', horizontal: 'center', };
                  // } else if (number == 2) {
                  //   cell.alignment = { indent: 1, vertical: 'middle', horizontal: 'center', };
                  // } else if (number > 1 && number < 7) {
                  //   cell.alignment = { indent: 1, vertical: 'middle', horizontal: 'center', };
                  // }
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
                if (d.NotesStatus == 'Y' && this.notesViewState == true) {
                  worksheet.mergeCells(notesCount, 1, notesCount, 31);
                  const Data2NOtes = worksheet.getCell(notesCount, 1);
                  Data2NOtes.value = 'Notes'
                  Data2NOtes.alignment = { indent: 2, vertical: 'middle', horizontal: 'left', };
                  Data2NOtes.font = { name: 'Arial', family: 4, size: 9 };

                  Data2NOtes.border = { right: { style: 'thin' }, left: { style: 'thin' }, top: { style: 'thin' }, bottom: { style: 'thin' } };
                  notesCount++

                  for (const d1 of d.Notes) {
                    worksheet.mergeCells(notesCount, 1, notesCount, 31);
                    const Data2 = worksheet.getCell(notesCount, 1);
                    Data2.value = ' ' + ' ' + ' ' + ' ' + ' ' + ' ' + d1.GN_Text
                    Data2.alignment = { indent: 2, vertical: 'middle', horizontal: 'left', };
                    Data2.font = { name: 'Arial', family: 4, size: 9 };
                    Data2.border = { right: { style: 'thin' }, left: { style: 'thin' }, top: { style: 'thin' }, bottom: { style: 'thin' } };
                    Data2.fill = {
                      type: 'pattern',
                      pattern: 'solid',
                      fgColor: { argb: 'e5e5e5' },
                      bgColor: { argb: 'FF0000FF' },
                    };
                    notesCount++
                  }
                }

              }
              worksheet.getColumn(1).width = 18;
              worksheet.getColumn(2).width = 20;
              worksheet.getColumn(3).width = 15;
              worksheet.getColumn(4).width = 30;
              worksheet.getColumn(5).width = 30;
              worksheet.getColumn(6).width = 30;
              worksheet.getColumn(7).width = 15;
              worksheet.getColumn(8).width = 30;
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
              worksheet.getColumn(22).width = 15;
              worksheet.getColumn(23).width = 15;
              worksheet.getColumn(24).width = 15;
              worksheet.getColumn(25).width = 15;
              worksheet.getColumn(26).width = 15;
              worksheet.addRow([]);
            
                this.shared.exportToExcel(workbook, 'Service Open RO Details' + DATE_EXTENSION );
             
            }
          }
        } else {
          this.shared.spinner.hide()
          this.toast.show(res.error,'danger','Error')
        }
      })
  }
}
