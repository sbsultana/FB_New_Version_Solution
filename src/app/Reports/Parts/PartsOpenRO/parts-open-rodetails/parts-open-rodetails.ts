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
  selector: 'app-parts-open-rodetails',
  imports: [SharedModule],
  templateUrl: './parts-open-rodetails.html',
  styleUrl: './parts-open-rodetails.scss',
})
export class PartsOpenRODetails {
  @Input() RODetailsObject: any = [];
  NoData!: boolean;
  spinnerLoader: boolean = true;
  spinnerLoadersec: boolean = false;
  pageNumber: any = 0;
  PartsPersonDetails: any = []
  constructor(public shared: Sharedservice, public setdates: Setdates, private comm: common, private cp: CurrencyPipe, private toast: ToastService, private ngbmodalActive: NgbActiveModal) {
  }
  ngOnInit(): void {
    //console.log(this.RODetailsObject);
    this.getDetails()

  }
  details: any = [];
  viewRO(roData: any) {
    // const modalRef = this.ngbmodel.open(RepairOrderComponent, { size: 'md', windowClass: 'compModal' });
    // modalRef.componentInstance.data = {  ro: roData.RONumber, storeid: roData.StoreID, vin:roData.vin, vehicleid:roData.vehicleid,source:roData.source, custno: roData.customernumber   }; // Pass data to the modal component    
    // modalRef.result.then((result) => {
    //   console.log(result); // Handle modal close result
    // }, (reason) => {
    //   console.log(`Dismissed: ${reason}`); // Handle dismiss reason
    // });
  }
  completeData: any = []

  getDetails() {
    // this.shared.spinner.show()
    const obj = {
      // "startdealdate": this.Partsdetails[0].StartDate,
      // "enddealdate":"10-17-2022",
      // "var1": "DealerName",
      // "var2": "AdvisorName",
      // "var3": "",
      // "var1Value": "Audi of Scottsdale",
      // "var2Value": "Shah Amir",
      // "var3Value": ""

      startdealdate: this.RODetailsObject[0].StartDate,
      enddealdate: this.RODetailsObject[0].EndDate,
      Labortype: this.RODetailsObject[0].Labortype,
      Saletype: this.RODetailsObject[0].Saletype,
      var1: this.RODetailsObject[0].var1.toString().indexOf(',') > 0 ? (this.RODetailsObject[0].type == 'C' ? this.RODetailsObject[0].var1.split(',')[1] : this.RODetailsObject[0].var1.split(',')[0]) : this.RODetailsObject[0].var1,
      var2: this.RODetailsObject[0].var2.toString().indexOf(',') > 0 ? (this.RODetailsObject[0].type == 'C' ? this.RODetailsObject[0].var2.split(',')[1] : this.RODetailsObject[0].var2.split(',')[0]) : this.RODetailsObject[0].var2,
      var3: '',
      var1Value: this.RODetailsObject[0].var1Value,
      var2Value: this.RODetailsObject[0].var2Value,
      var3Value: this.RODetailsObject[0].var3Value,
      PageNumber: this.pageNumber,
      PageSize: '100',
      minage: this.RODetailsObject[0].minage,
      maxage: this.RODetailsObject[0].maxage,
      Oldro: ""
    };
    this.shared.api
      .postmethod(this.comm.routeEndpoint + 'GetPartsGrossSummaryDetailsOpen', obj)
      .subscribe((res) => {
        if (res.status == 200) {
          this.details = res.response;
          this.completeData = [
            ...this.PartsPersonDetails,
            ...this.details,
          ]
          this.PartsPersonDetails = this.completeData

          // this.PartsPersonDetails = [
          //   ...this.PartsPersonDetails,
          //   ...this.details,
          // ];
          //console.log(this.PartsPersonDetails);
          // this.shared.spinner.hide()
          this.spinnerLoader = false;
          this.spinnerLoadersec = false;
          // this.PartsPersonDetails=res.response
          //console.log(this.PartsPersonDetails);
          // this.shared.spinner.hide()
          this.spinnerLoader = false;
          if (this.PartsPersonDetails.length > 0) {
            this.NoData = false;
          } else {
            this.NoData = true;
          }
        } else {
          this.spinnerLoader = false;
          this.NoData = true;
        }
      }, (error) => {
        this.spinnerLoader = false;
        this.NoData = true;
      }

      );
  }

  isDesc: boolean = false;
  column: string = 'CategoryName';
  notesViewState: boolean = false
  sort(property: any, data: any) {
    this.isDesc = !this.isDesc; //change the direction
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
  @ViewChild('scrollOne') scrollOne!: ElementRef;
  updateVerticalScroll(event: any): void {

    this.scrollOne.nativeElement.scrollTop = event.target.scrollTop;
    if (
      event.target.scrollTop + event.target.clientHeight >=
      event.target.scrollHeight - 2
    ) {
      // this.toast.show("reached at bottom");
      if (this.pageNumber == 0) {
        if (this.details.length == 100) {
          this.spinnerLoader = true;
          this.pageNumber++;
          this.getDetails();
        }
      } else {
        if (this.details.length >= 100) {
          this.spinnerLoader = true;
          this.pageNumber++;
          this.getDetails();
        }
      }
    }

  }
  comments: any = []
  popupOpen: any = ''
  openCommentsPopUp(item: any, temp: any) {
    this.comments = [
      {

        var1: this.RODetailsObject[0].var2Value == '' || this.RODetailsObject[0].var2Value == undefined ? this.RODetailsObject[0].var1Value : this.RODetailsObject[0].var2Value,
        moduleId: '83',
        moduleName: 'SOR',
        type: item.ronumber

        // var3Value: Item.data3,
        // userName: Item.data3,
      },

    ];
    this.ngbmodalActive = this.shared.ngbmodal.open(
      temp,
      {
        size: 'xxl',
        backdrop: 'static',
      }
    );

    // DetailsSalesPeron.result.then(
    //   (data) => { },
    //   (reason) => {
    //     // on dismiss
    //     // this.getAllComments(item.ronumber,item,item.commet=='-'?'show':'hide')
    //     this.pageNumber=0
    //     this.ServiceData=[];
    //     this.spinnerLoader= true;
    //     this.getDetails()

    //   }
    // );
  }
  QISearchName: any = ''
  filteredServiceOpenROData: any = []
  filterData() {
    if (this.QISearchName.trim() !== '') {
      this.PartsPersonDetails = this.completeData.filter((item: any) =>
        (item.RONumber && item.RONumber.toLowerCase().includes(this.QISearchName.toLowerCase())) ||
        (item.CDate && item.CDate.toLowerCase().includes(this.QISearchName.toLowerCase())) ||
        (typeof item.Age === 'string' && item.Age.toLowerCase().includes(this.QISearchName.toLowerCase())) ||
        (item.Customername && item.Customername.toLowerCase().includes(this.QISearchName.toLowerCase())) ||
        (item.InvoiceNumber && item.InvoiceNumber.toLowerCase().includes(this.QISearchName.toLowerCase())) ||
        (item.RONumber && item.RONumber.toLowerCase().includes(this.QISearchName.toLowerCase())) ||
        (item.AP_description && item.AP_description.toLowerCase().includes(this.QISearchName.toLowerCase()))
      );
    } else {
      this.PartsPersonDetails = [...this.completeData];
    }
    // this.callLoadingState == 'ANS' ? this.sort(this.column, this.filteredServiceOpenROData, this.callLoadingState) : ''
    // let position = this.scrollpositionstoring + 10
    // setTimeout(() => {
    //   this.scrollcent.nativeElement.scrollTop = position
    //   // //console.log(position);

    // }, 500);
    // this.pageNumber = 1;
  }

  closePopup(event: any) {
    this.pageNumber = 0
    this.PartsPersonDetails = [];
    this.popupOpen = 'Y'

    this.spinnerLoader = true;
    this.getDetails()
  }
  commentslist: any = []
  getAllComments(type: any, item: any, condition: any) {
    if (condition == 'show') {
      item.comment = '-'
      this.commentslist = [];
      // item.commentslist = [];
      this.spinnerLoader = true;
      //console.log(localStorage.getItem('UserDetails'));
      let ud = localStorage.getItem('UserDetails');
      const obj = {
        UserId: JSON.parse(ud!).userid,
        ModuleName: 'SOR',
        Title: type,
      };
      //console.log(obj);
      this.shared.api.postmethod('whispercomments/get', obj).subscribe(
        (res) => {
          if (res.status == 200) {
            this.spinnerLoader = false;
            this.commentslist = res.response;
            item.commentslist = res.response
            item.comment = '-'
            //console.log(item);

          } else {
            this.toast.show('Invalid Details','danger','Error');
          }
        },
        (error) => {
          // // //console.log(error);
        }
      );
    }
    if (condition == 'hide') {
      item.commentslist = []
      item.comment = '+'
    }

  }
  commentsVisibility: boolean = true
  openComments() {
    this.commentsVisibility = !this.commentsVisibility
  }
async openServiceModal(roNumber: any, vin: any, storeid: any, vehicleid: any, source: any, custno: any) {
    const module = await import('../../../../Layout/cdpdataview/repair/repair-module');
    const component = module.Repair;
    const modalRef = this.shared.ngbmodal.open(component, { size: 'xl', windowClass: 'compModal' });
    modalRef.componentInstance.data = { ro: roNumber, vin: vin, storeid: storeid, vehicleid: vehicleid, source: source, custno: custno }; // Pass data to the modal component
    modalRef.result.then((result) => {
      console.log(result); // Handle modal close result
    }, (reason) => {
      console.log(`Dismissed: ${reason}`); // Handle dismiss reason
    });
  }
  exportAsXLSX() {
    this.shared.spinner.show()
    const obj = {
      startdealdate: this.RODetailsObject[0].StartDate,
      enddealdate: this.RODetailsObject[0].EndDate,
      Labortype: this.RODetailsObject[0].Labortype,
      Saletype: this.RODetailsObject[0].Saletype,
      var1: this.RODetailsObject[0].var1.toString().indexOf(',') > 0 ? (this.RODetailsObject[0].type == 'C' ? this.RODetailsObject[0].var1.split(',')[1] : this.RODetailsObject[0].var1.split(',')[0]) : this.RODetailsObject[0].var1,
      var2: this.RODetailsObject[0].var2.toString().indexOf(',') > 0 ? (this.RODetailsObject[0].type == 'C' ? this.RODetailsObject[0].var2.split(',')[1] : this.RODetailsObject[0].var2.split(',')[0]) : this.RODetailsObject[0].var2,
      var3: '',
      var1Value: this.RODetailsObject[0].var1Value,
      var2Value: this.RODetailsObject[0].var2Value,
      var3Value: this.RODetailsObject[0].var3Value,
      PageNumber: '0',
      PageSize: '10000',
      minage: this.RODetailsObject[0].minage,
      maxage: this.RODetailsObject[0].maxage,
    };
    this.shared.api
      .postmethod(this.comm.routeEndpoint + 'GetPartsGrossSummaryDetailsOpen', obj).subscribe(
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
                const worksheet = workbook.addWorksheet('Parts Open RO Details');
                worksheet.views = [
                  {
                    state: 'frozen',
                    ySplit: 9, // Number of rows to freeze (2 means the first two rows are frozen)
                    topLeftCell: 'A10', // Specify the cell to start freezing from (in this case, the third row)
                    showGridLines: false,
                  },
                ];
                const DateToday = this.shared.datePipe.transform(new Date(), 'MM/dd/yyyy h:mm:ss a');

                const titleRow = worksheet.getCell("A2"); titleRow.value = 'Parts Open RO Details';
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
                  'Open Date',
                  'Customer',
                  'Invoice',
                  'Age',
                  'Part #',
                  'RO Number',
                  'Description',

                  'Total Gross',
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

                  'Internal Gross',
                  'Internal GP%',
                  'Internal ELR',
                  'Internal Hours',
                  'Internal Discount',

                  'Counter Retail Gross',
                  'Counter Retail GP%',
                  'Counter Retail ELR',
                  'Counter Retail Hours',
                  'Counter Retail Discount',

                  'Wholesale Gross',
                  'Wholesale GP%',
                  'Wholesale ELR',
                  'Wholesale Hours',
                  'Wholesale Discount',
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
                  d.opendate = this.shared.datePipe.transform(d.CDate, 'MM/dd/yyyy');
                  const Data1 = worksheet.addRow([

                    d.opendate == '' ? '-' : d.opendate == null ? '-' : d.opendate,
                    d.Customername == '' ? '-' : d.Customername == null ? '-' : d.Customername,
                    d.InvoiceNumber == '' ? '-' : d.InvoiceNumber == null ? '-' : d.InvoiceNumber.toString(),
                    d.Age == '' ? '-' : d.Age == null ? '-' : d.Age.toString(),
                    d.Partnumber == '' ? '-' : d.Partnumber == null ? '-' : d.Partnumber.toString(),
                    d.RONumber == '' ? '-' : d.RONumber == null ? '-' : d.RONumber.toString(),
                    d.AP_description == '' ? '-' : d.AP_description == null ? '-' : d.AP_description.toString(),

                    d.Total_Gross == '' ? '-' : d.Total_Gross == null ? '-' : parseFloat(d.Total_Gross),
                    d.Total_Retention == '' ? '-' : d.Total_Retention == null ? '-' : parseFloat(d.Total_Retention) + ' %',
                    d.Total_ELR == '' ? '-' : d.Total_ELR == null ? '-' : parseFloat(d.Total_ELR),
                    d.Total_Hours == '' ? '-' : d.Total_Hours == null ? '-' : parseFloat(d.Total_Hours),
                    d.Total_Discount == '' ? '-' : d.Total_Discount == null ? '-' : parseFloat(d.Total_Discount),

                    d.Cust_Gross == '' ? '-' : d.Cust_Gross == null ? '-' : parseFloat(d.Cust_Gross),
                    d.Cust_Retention == '' ? '-' : d.Cust_Retention == null ? '-' : parseFloat(d.Cust_Retention) + ' %',
                    d.Cust_ELR == '' ? '-' : d.Cust_ELR == null ? '-' : parseFloat(d.Cust_ELR),
                    d.Cust_Hours == '' ? '-' : d.Cust_Hours == null ? '-' : parseFloat(d.Cust_Hours),
                    d.Cust_Discount == '' ? '-' : d.Cust_Discount == null ? '-' : parseFloat(d.Cust_Discount),

                    d.Warranty_Gross == '' ? '-' : d.Warranty_Gross == null ? '-' : parseFloat(d.Warranty_Gross),
                    d.Warranty_Retention == '' ? '-' : d.Warranty_Retention == null ? '-' : parseFloat(d.Warranty_Retention) + ' %',
                    d.Warranty_ELR == '' ? '-' : d.Warranty_ELR == null ? '-' : parseFloat(d.Warranty_ELR),
                    d.Warranty_Hours == '' ? '-' : d.Warranty_Hours == null ? '-' : parseFloat(d.Warranty_Hours),
                    d.Warranty_Discount == '' ? '-' : d.Warranty_Discount == null ? '-' : parseFloat(d.Warranty_Discount),

                    d.Internal_Gross == '' ? '-' : d.Internal_Gross == null ? '-' : parseFloat(d.Internal_Gross),
                    d.Internal_Retention == '' ? '-' : d.Internal_Retention == null ? '-' : parseFloat(d.Internal_Retention) + ' %',
                    d.Internal_ELR == '' ? '-' : d.Internal_ELR == null ? '-' : parseFloat(d.Internal_ELR),
                    d.Internal_Hours == '' ? '-' : d.Internal_Hours == null ? '-' : parseFloat(d.Internal_Hours),
                    d.Internal_Discount == '' ? '-' : d.Internal_Discount == null ? '-' : parseFloat(d.Internal_Discount),

                    d.Counter_Gross == '' ? '-' : d.Counter_Gross == null ? '-' : parseFloat(d.Counter_Gross),
                    d.Counter_Retention == '' ? '-' : d.Counter_Retention == null ? '-' : parseFloat(d.Counter_Retention) + ' %',
                    d.Counter_ELR == '' ? '-' : d.Counter_ELR == null ? '-' : parseFloat(d.Counter_ELR),
                    d.Counter_Hours == '' ? '-' : d.Counter_Hours == null ? '-' : parseFloat(d.Counter_Hours),
                    d.Counter_Discount == '' ? '-' : d.Counter_Discount == null ? '-' : parseFloat(d.Counter_Discount),

                    d.Wholesale_Gross == '' ? '-' : d.Wholesale_Gross == null ? '-' : parseFloat(d.Wholesale_Gross),
                    d.Wholesale_Retention == '' ? '-' : d.Wholesale_Retention == null ? '-' : parseFloat(d.Wholesale_Retention) + ' %',
                    d.Wholesale_ELR == '' ? '-' : d.Wholesale_ELR == null ? '-' : parseFloat(d.Wholesale_ELR),
                    d.Wholesale_Hours == '' ? '-' : d.Wholesale_Hours == null ? '-' : parseFloat(d.Wholesale_Hours),
                    d.Wholesale_Discount == '' ? '-' : d.Wholesale_Discount == null ? '-' : parseFloat(d.Wholesale_Discount),
                  ]);
                  // Data1.outlineLevel = 1; // Grouping level 1
                  Data1.font = { name: 'Arial', family: 4, size: 8 };
                  Data1.height = 18;
                  // Data1.getCell(1).alignment = {indent: 1,vertical: 'middle', horizontal: 'left'}
                  Data1.eachCell((cell, number) => {
                    cell.border = { right: { style: 'thin' } };
                    cell.numFmt = '$#,##0.00';
                    if (number == 11 || number == 16 || number == 21 || number == 26 || number == 31 || number == 36) {
                      cell.numFmt = '0.00';
                      cell.alignment = { indent: 1, vertical: 'middle', horizontal: 'center', };
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
                  if (d.NotesStatus == 'Y' && this.notesViewState == true) {
                    worksheet.mergeCells(notesCount, 1, notesCount, 29);
                    const Data2NOtes = worksheet.getCell(notesCount, 1);
                    Data2NOtes.value = 'Notes'
                    Data2NOtes.alignment = { indent: 2, vertical: 'middle', horizontal: 'left', };
                    Data2NOtes.font = { name: 'Arial', family: 4, size: 9 };

                    Data2NOtes.border = { right: { style: 'thin' }, left: { style: 'thin' }, top: { style: 'thin' }, bottom: { style: 'thin' } };
                    notesCount++

                    for (const d1 of d.Notes) {
                      worksheet.mergeCells(notesCount, 1, notesCount, 37);
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
              
                  this.shared.exportToExcel(workbook, 'Parts Open RO Details' + DATE_EXTENSION );
                
              }
            }
          } else {
            this.shared.spinner.hide()
            this.toast.show(res.error,'danger','Error')
          }
        })
  }

}