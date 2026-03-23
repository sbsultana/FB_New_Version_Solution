import { Component, ElementRef, HostListener, Input, OnInit, Renderer2, ViewChild, } from '@angular/core';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { FilterPipe, FilterPipeModule } from 'ngx-filter-pipe';


@Component({
  selector: 'app-servicegross-details',
  imports: [SharedModule, FilterPipeModule],
  providers: [FilterPipe],
  templateUrl: './parts-gross-details.html',
  styleUrl: './parts-gross-details.scss',
})
export class PartsGrossDetails {
  @Input() Partsdetails: any = [];
  PartsPersonDetails: any = [];
  NoData!: boolean;
  spinnerLoader: boolean = true;

  pageNumber: any = 0;
  spinnerLoadersec: boolean = false;
  details: any = [];


  @HostListener('scroll', ['$event'])
  onScroll(event: Event): void {
    const target = event.target as HTMLElement;

    // Check if the scroll position is at the bottom
    const isAtBottom = target.scrollHeight - target.scrollTop === target.clientHeight;

    if (isAtBottom) {
      //console.log('Scroll position is at the bottom');
    }
  }
  constructor(    public shared: Sharedservice,    private renderer: Renderer2,  ) {
    this.renderer.listen('window', 'click', (e: Event) => {
      const TagName = e.target as HTMLButtonElement;
      if (TagName.className === 'd-block modal fade show modal-static') {
        // this.closeBtn.nativeElement.click();
        this.close();
      }
    });

  }

  ngOnInit() {
    // this.spinnerLoader=false
    //console.log(this.Partsdetails);
    this.GetDetails();
  }
  GetDetails() {
    // this.shared.spinner.show()
    const obj = {


      startdealdate: this.Partsdetails[0].StartDate,
      enddealdate: this.Partsdetails[0].EndDate,
      Store: this.Partsdetails[0].Store == undefined ? '' : this.Partsdetails[0].Store,
      Labortype: this.Partsdetails[0].Labortype,
      Saletype: this.Partsdetails[0].Saletype,
      // PartsSource: this.Partsdetails[0].PartsSource,
      SourceBulk: this.Partsdetails[0].PartsSource == 'B' ? 'Y' : '',
      SourceTire: this.Partsdetails[0].PartsSource == 'T' ? 'Y' : '',
      SourceWithout: this.Partsdetails[0].PartsSource == 'N' ? 'Y' : '',
      var1: this.Partsdetails[0].var1,
      var2: this.Partsdetails[0].var2,
      var3: '',
      var1Value: this.Partsdetails[0].var1Value,
      var2Value: this.Partsdetails[0].var2Value,
      var3Value: this.Partsdetails[0].var3Value,
      PageNumber: this.pageNumber,
      PageSize: '100',
    };
   this.shared.api
      .postmethod(this.shared.common.routeEndpoint + 'GetPartsGrossSummaryDetailsNew', obj)
      .subscribe((res) => {
        if (res.status == 200) {
          this.details = res.response;
          this.PartsPersonDetails = [
            ...this.PartsPersonDetails,
            ...this.details,
          ];
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
            let position = this.scrollCurrentposition + 10
            setTimeout(() => {
              this.scrollOne.nativeElement.scrollTop = position
              this.scrollTwo.nativeElement.scrollTop = position

            }, 500);

            // //console.log(this.scrollCurrentposition,this.scrollTwo.nativeElement.scrollTop,position);

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

  getTotal(columnname: any) {
    let total = 0
    // return marks.reduce((acc: number, {InStock}: any) => acc += +(InStock || 0), 0);
    // this.InventorySummaryData[0].data2.forEach((val: any, i: any) => {
    console.log(columnname);

    this.PartsPersonDetails.some(function (x: any) {
      // if(columnname == 'AvgMonth'){
      //   console.log(parseInt(x[columnname]));
      // total += x[columnname]        
      // }else{
      total += x[columnname]

      // }
    })
    return total
    // })
    // frontgross.some(function (x: any) {
    //   total += parseInt(x[colname])
    // })

  }
  close() {
    this.shared.ngbmodal.dismissAll();
  }

  currentElement!: string;

  @ViewChild('scrollOne') scrollOne!: ElementRef;
  @ViewChild('scrollTwo') scrollTwo!: ElementRef;
  count = 0
  scrollCurrentposition: any = 0
  updateVerticalScroll(event: any): void {
    this.scrollCurrentposition = event.scrollTop
    // //console.log(this.scrollCurrentposition,event);

    if (this.currentElement === 'scrollTwo') {
      this.scrollOne.nativeElement.scrollTop = event.scrollTop;

      if (
        event.scrollTop + event.clientHeight >=
        event.scrollHeight - 2
      ) {
        this.count++
        //console.log(event.scrollTop + event.clientHeight,
        // event.scrollHeight, this.count);

        // alert("reached at bottom");
        if (this.count % 2 == 0) {
          if (this.pageNumber == 0) {
            if (this.details.length == 100) {
              this.spinnerLoadersec = true;
              this.pageNumber++;
              this.GetDetails();
            }
            // this.spinnerLoadersec = true;
            // this.pageNumber++;
            // this.GetDetails();
          } else {
            if (this.details.length >= 100) {
              this.spinnerLoadersec = true;
              this.pageNumber++;
              this.GetDetails();
            }
          }
        }

      }
    }
    else if (this.currentElement === 'scrollOne') {
      this.scrollTwo.nativeElement.scrollTop = event.scrollTop;
      if (
        event.scrollTop + event.clientHeight >=
        event.scrollHeight
      ) {
        this.count++
        if (this.count % 2 == 0) {
          if (this.pageNumber == 0) {
            if (this.details.length == 100) {
              this.spinnerLoadersec = true;
              this.pageNumber++;
              this.GetDetails();
            }
            // this.spinnerLoadersec = true;
            // this.pageNumber++;
            // this.GetDetails();
          } else {
            if (this.details.length >= 100) {
              this.spinnerLoadersec = true;
              this.pageNumber++;
              this.GetDetails();
            }
          }
        }
      }
    }
  }

  trTag: any = '';
  secondtrtag: any = ''

  hoverclass(e: any, i: any) {
    if (this.trTag != '') {
      this.trTag.classList.remove('hover');
      this.secondtrtag.classList.remove('hover');
    }
    let id = (e.target as Element).id;
    this.trTag = document.getElementById('SD_' + i);
    this.secondtrtag = document.getElementById('SP_' + i);
    if (id == 'SD_' + i || id == 'SP_' + i) {
      this.trTag.classList.add('hover');
      this.secondtrtag.classList.add('hover');
    }
  }
  updateCurrentElement(element: 'scrollOne' | 'scrollTwo') {
    this.currentElement = element;
  }

    public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    }
    return false;
  }
  exportAsXLSX() {
    this.shared.spinner.show()
    const obj = {


      startdealdate: this.Partsdetails[0].StartDate,
      enddealdate: this.Partsdetails[0].EndDate,
      Store: this.Partsdetails[0].Store == undefined ? '' : this.Partsdetails[0].Store,
      Labortype: this.Partsdetails[0].Labortype,
      Saletype: this.Partsdetails[0].Saletype,
      // PartsSource: this.Partsdetails[0].PartsSource,
      SourceBulk: this.Partsdetails[0].PartsSource == 'B' ? 'Y' : '',
      SourceTire: this.Partsdetails[0].PartsSource == 'T' ? 'Y' : '',
      SourceWithout: this.Partsdetails[0].PartsSource == 'N' ? 'Y' : '',
      var1: this.Partsdetails[0].var1,
      var2: this.Partsdetails[0].var2,
      var3: '',
      var1Value: this.Partsdetails[0].var1Value,
      var2Value: this.Partsdetails[0].var2Value,
      var3Value: this.Partsdetails[0].var3Value,
      PageNumber: 0,
      PageSize: '10000',
    };
   this.shared.api
      .postmethod(this.shared.common.routeEndpoint + 'GetPartsGrossSummaryDetailsNew', obj)
      .subscribe((res) => {
        if (res.status == 200) {
          this.shared.spinner.hide()

          if (res.response != undefined) {
            if (res.response.length > 0) {


              let localarray = res.response.map((_arrayElement: any) =>
                Object.assign({}, _arrayElement)
              );
              const workbook = this.shared.getWorkbook();
              const worksheet = workbook.addWorksheet('Parts Gross Details');
              worksheet.views = [
                {
                  state: 'frozen',
                  ySplit: 9, // Number of rows to freeze (2 means the first two rows are frozen)
                  topLeftCell: 'A10', // Specify the cell to start freezing from (in this case, the third row)
                  showGridLines: false,
                },
              ];
              const DateToday = this.shared.datePipe.transform(new Date(), 'MM/dd/yyyy h:mm:ss a');

              const titleRow = worksheet.getCell("A2"); titleRow.value = 'Parts Gross Details';
              titleRow.font = { name: 'Arial', family: 4, size: 15, bold: true };
              titleRow.alignment = { indent: 1, vertical: 'middle', horizontal: 'left' }



              const DateBlock = worksheet.getCell("L2"); DateBlock.value = DateToday;
              DateBlock.font = { name: 'Arial', family: 4, size: 10 };
              DateBlock.alignment = { vertical: 'middle', horizontal: 'center' }
              worksheet.addRow([''])
              const Store_Name = worksheet.addRow(['Store Name :']);
              Store_Name.getCell(1).font = { name: 'Arial', family: 4, size: 9, bold: true, };
              Store_Name.alignment = { indent: 1, vertical: 'middle', horizontal: 'left' }
              const StoreName = worksheet.getCell("B4"); StoreName.value = this.Partsdetails[0].var1Value;
              StoreName.font = { name: 'Arial', family: 4, size: 9 };
              StoreName.alignment = { indent: 1, vertical: 'middle', horizontal: 'left' }

              const DATE_EXTENSION = this.shared.datePipe.transform(
                new Date(),
                'MMddyyyy'
              );

              const StartDealDate = worksheet.addRow(['Start Date :']);
              const startdealdate = worksheet.getCell('B5');
              startdealdate.value = this.Partsdetails[0].StartDate;
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
              enddealdate.value = this.Partsdetails[0].EndDate;
              enddealdate.font = { name: 'Arial', family: 4, size: 9 };
              enddealdate.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
              EndDealDate.getCell(1).font = {
                name: 'Arial',
                family: 4,
                size: 9,
                bold: true,
              };

              const Var1Value = worksheet.getCell("E7")
              Var1Value.value = 'Details of:' + (this.Partsdetails[0].userName == '' ? '-' : this.Partsdetails[0].userName == null ? '-' : this.Partsdetails[0].userName == undefined ? '--' : (this.Partsdetails[0].userName == 'C' ?
                'Customer Pay':(this.Partsdetails[0].userName == 'T' ? 'Warranty' : (this.Partsdetails[0].userName
                == 'I' ? 'Internal' : (this.Partsdetails[0].userName == 'R' ? 'Retail':
                (this.Partsdetails[0].userName == 'W' ? 'Wholesale': ( this.Partsdetails[0].userName == 'E' ? 'Entended Warranty' : (this.Partsdetails[0].userName))))))));
              Var1Value.font = { name: 'Arial', family: 4, size: 11, bold: true }
              Var1Value.alignment = { indent: 1, vertical: 'middle', horizontal: 'left' }
              worksheet.addRow('');
              let Headings = [
                'Sl.no',
                'Invoice #',
                'Date',
                'Customer',
                'Part #',
                'Total Sale',
                'Total Gross',
                'Total GP%',
                'Customer Pay Sales',
                'Customer Pay Gross',
                'Customer Pay GP%',
                'Warranty Sales',
                'Warranty Gross',
                'Warranty GP%',
                'Extended Warranty Sales',
                'Extended Warranty Gross',
                'Extended Warranty GP%',
                'Internal Sales',
                'Internal Gross',
                'Internal GP%',
                'Counter Retail Sales',
                'Counter Retail Gross',
                'Counter Retail GP%',
                'Wholesale Sales',
                'Wholesale Gross',
                'Wholesale GP%',
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
              for (const d of localarray) {
                count++
                d.CDate = this.shared.datePipe.transform(d.CDate, 'MM.dd.yyyy');
                const Data1 = worksheet.addRow([
                  count,
                  d.InvoiceNumber == '' ? '-' : d.InvoiceNumber == null ? '-' : d.InvoiceNumber,
                  d.CDate == '' ? '-' : d.CDate == null ? '-' : d.CDate,
                  d.Customername == '' ? '-' : d.Customername == null ? '-' : d.Customername,
                  d.Partnumber == '' ? '-' : d.Partnumber == null ? '-' : d.Partnumber,
                  d.Total_Sale == '' ? '-' : d.Total_Sale == null ? '-' : parseFloat(d.Total_Sale),
                  d.Total_Gross == '' ? '-' : d.Total_Gross == null ? '-' : parseFloat(d.Total_Gross),
                  d.Total_Retention == '' ? '-' : d.Total_Retention == null ? '-' : d.Total_Retention + '%',
                  d.Cust_Sale == '' ? '-' : d.Cust_Sale == null ? '-' : parseFloat(d.Cust_Sale),

                  d.Cust_Gross == '' ? '-' : d.Cust_Gross == null ? '-' : parseFloat(d.Cust_Gross),
                  d.Cust_Retention == '' ? '-' : d.Cust_Retention == null ? '-' : d.Cust_Retention + '%',

                  d.Warranty_Sale == '' ? '-' : d.Warranty_Sale == null ? '-' : parseFloat(d.Warranty_Sale),
                  d.Warranty_Gross == '' ? '-' : d.Warranty_Gross == null ? '-' : parseFloat(d.Warranty_Gross),
                  d.Warranty_Retention == '' ? '-' : d.Warranty_Retention == null ? '-' : d.Warranty_Retention + '%',

                  d.Extended_Sale == '' ? '-' : d.Extended_Sale == null ? '-' : parseFloat(d.Extended_Sale),
                  d.Extended_Gross == '' ? '-' : d.Extended_Gross == null ? '-' : parseFloat(d.Extended_Gross),
                  d.Extended_Retention == '' ? '-' : d.Extended_Retention == null ? '-' : d.Extended_Retention + '%',

                  d.Internal_Sale == '' ? '-' : d.Internal_Sale == null ? '-' : parseFloat(d.Internal_Sale),
                  d.Internal_Gross == '' ? '-' : d.Internal_Gross == null ? '-' : parseFloat(d.Internal_Gross),
                  d.Internal_Retention == '' ? '-' : d.Internal_Retention == null ? '-' : d.Internal_Retention + '%',
                  d.Counter_sale == '' ? '-' : d.Counter_sale == null ? '-' : parseFloat(d.Counter_sale),

                  d.Counter_Gross == '' ? '-' : d.Counter_Gross == null ? '-' : parseFloat(d.Counter_Gross),
                  d.Counter_Retention == '' ? '-' : d.Counter_Retention == null ? '-' : d.Counter_Retention + '%',
                  d.Wholesale_Sale == '' ? '-' : d.Wholesale_Sale == null ? '-' : parseFloat(d.Wholesale_Sale),
                  d.Wholesale_Gross == '' ? '-' : d.Wholesale_Gross == null ? '-' : parseFloat(d.Wholesale_Gross),
                  d.Wholesale_Retention == '' ? '-' : d.Wholesale_Retention == null ? '-' : d.Wholesale_Retention + '%',

                ]);
                // Data1.outlineLevel = 1; // Grouping level 1
                Data1.font = { name: 'Arial', family: 4, size: 8 };
                Data1.height = 18;
                Data1.alignment = { indent: 1, vertical: 'middle', horizontal: 'center' }
                Data1.eachCell((cell, number) => {
                  cell.border = { right: { style: 'thin' } };
                  cell.numFmt = '$#,##0.00';

                  if (number > 6 && number < 27) {
                    cell.alignment = { indent: 1, vertical: 'middle', horizontal: 'center', };
                  } else if (number == 1) {
                    cell.numFmt = '###0';
                    cell.alignment = { indent: 1, vertical: 'middle', horizontal: 'center', };
                  } else if (number == 2) {
                    cell.alignment = { indent: 1, vertical: 'middle', horizontal: 'center', };
                  } else if (number > 1 && number < 7) {
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
              }
              worksheet.getColumn(1).width = 18;
              worksheet.getColumn(2).width = 20;
              worksheet.getColumn(3).width = 15;
              worksheet.getColumn(4).width = 30;
              worksheet.getColumn(5).width = 30;
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
              worksheet.getColumn(22).width = 15;
              worksheet.getColumn(23).width = 15;
              worksheet.getColumn(24).width = 15;
              worksheet.getColumn(25).width = 15;
              worksheet.getColumn(26).width = 15;
              worksheet.getColumn(27).width = 15;
              worksheet.getColumn(28).width = 15;
              worksheet.getColumn(29).width = 15;
              worksheet.getColumn(30).width = 15;
              worksheet.getColumn(31).width = 15;
              worksheet.getColumn(32).width = 15;
              worksheet.getColumn(33).width = 15;
              worksheet.getColumn(34).width = 15;
              worksheet.getColumn(35).width = 15;
              worksheet.addRow([]);
              workbook.xlsx.writeBuffer().then((data: any) => {
                const blob = new Blob([data], {
                  type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                });
              this.shared.exportToExcel(workbook, 'Parts Gross Details' + DATE_EXTENSION );
              });
            }
          }
        }
        else {
          this.shared.spinner.hide()
          // this.toast.show(res.error,'danger','Error')
        }
      })
  }
}
