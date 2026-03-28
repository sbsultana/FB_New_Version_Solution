import { Component, ElementRef, Input, OnInit, Renderer2, ViewChild, } from '@angular/core';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { FilterPipe, FilterPipeModule } from 'ngx-filter-pipe';


@Component({
  selector: 'app-servicegross-details',
  imports: [SharedModule, FilterPipeModule],
  providers: [FilterPipe],
  templateUrl: './servicegross-details.html',
  styleUrl: './servicegross-details.scss',
})
export class ServicegrossDetails {
  @Input() Servicedetails: any = [];
  ServicePersonDetails: any = [];
  NoData!: boolean;
  spinnerLoader: boolean = true;
  spinnerLoadersec: boolean = false;
  pageNumber: any = 0;

  constructor(
    public shared: Sharedservice,
    private renderer: Renderer2,
  ) {
    this.renderer.listen('window', 'click', (e: Event) => {
      const TagName = e.target as HTMLButtonElement;
      if (TagName.className === 'd-block modal fade show modal-static') {
        // this.closeBtn.nativeElement.click();
        this.close();
      }
    });
  }
  viewRO(roData: any) {
    // const modalRef = this.ngbmodel.open(RepairOrderComponent, { size: 'md', windowClass: 'connectedmodal' });
    // modalRef.componentInstance.data = {  ro: roData.ronumber, storeid: roData.StoreID, vin:roData.vin, vehicleid:roData.vehicleid ,custno: roData?.customernumber  }; // Pass data to the modal component    
    // modalRef.result.then((result) => {
    //   console.log(result); // Handle modal close result
    // }, (reason) => {
    //   console.log(`Dismissed: ${reason}`); // Handle dismiss reason
    // });
  }
  ngOnInit() {
    // this.spinnerLoader=false
    //console.log(this.Servicedetails);
    this.GetDetails();
  }
  details: any = [];
  GetDetails() {
    // this.shared.spinner.show()
    const obj = {
      startdealdate: this.Servicedetails[0].StartDate,
      enddealdate: this.Servicedetails[0].EndDate,
      var1: this.Servicedetails[0].var1,
      var2: this.Servicedetails[0].var2,
      var3: '',
      var1Value: this.Servicedetails[0].var1Value,
      var2Value: this.Servicedetails[0].var2Value,
      var3Value: '',
      PaytypeC: this.Servicedetails[0].PaytypeC,
      PaytypeW: this.Servicedetails[0].PaytypeW,
      PaytypeI: this.Servicedetails[0].PaytypeI,
      DepartmentS: this.Servicedetails[0].DepartmentS,
      DepartmentP: this.Servicedetails[0].DepartmentP,
      DepartmentQ: this.Servicedetails[0].DepartmentQ,
      DepartmentB: this.Servicedetails[0].DepartmentB,
      PolicyAccount: this.Servicedetails[0].PolicyAccount,
      excludeZeroHours: this.Servicedetails[0].zeroHours,
      PageNumber: this.pageNumber,
      PageSize: '100',
      // "startdealdate":this.Servicedetails[0].StartDate,
      // "enddealdate":this.Servicedetails[0].EndDate,
      // "var1": "Store_Name",
      // "var2": "ServiceAdvisor_Name",
      // "var3": "",
      // "var1Value":this.Servicedetails[0].var1Value,
      // "var2Value":this.Servicedetails[0].var2Value,
      // "var3Value":this.Servicedetails[0].var3Value
    };
    this.shared.api
      .postmethod(this.shared.common.routeEndpoint + 'GetServicesGrossSummaryDetailsV2', obj)
      .subscribe((res) => {
        if (res.status == 200) {
          this.details = res.response;
          this.ServicePersonDetails = [
            ...this.ServicePersonDetails,
            ...this.details,
          ];
          // //console.log(this.ServicePersonDetails);
          // this.shared.spinner.hide()
          this.spinnerLoader = false;
          this.spinnerLoadersec = false;

          if (this.ServicePersonDetails.length > 0) {
            this.NoData = false;
          } else {
            this.NoData = true;
          }
        }
      });
  }
  close() {
    this.shared.ngbmodal.dismissAll();
  }

  currentElement!: string;

  @ViewChild('scrollOne') scrollOne!: ElementRef;
  @ViewChild('scrollTwo') scrollTwo!: ElementRef;
  count = 0
  updateVerticalScroll(event: any): void {
    if (this.currentElement === 'scrollTwo') {
      this.scrollOne.nativeElement.scrollTop = event.target.scrollTop;
      if (
        event.target.scrollTop + event.target.clientHeight >=
        event.target.scrollHeight - 2
      ) {
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
      this.scrollTwo.nativeElement.scrollTop = event.target.scrollTop;
      if (
        event.target.scrollTop + event.target.clientHeight >=
        event.target.scrollHeight - 2
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

  updateCurrentElement(element: 'scrollOne' | 'scrollTwo') {
    this.currentElement = element;
  }


  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    }
    return false;
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
  async openServiceModal(roNumber: any, vin: any, storeid: any, vehicleid: any, source: any, custno: any) {
    const module = await import('../../../../Layout/cdpdataview/repair/repair-module');
    const component = module.Repair;
    const modalRef = this.shared.ngbmodal.open(component, { size: 'xl', windowClass: 'connectedmodal' });
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
      startdealdate: this.Servicedetails[0].StartDate,
      enddealdate: this.Servicedetails[0].EndDate,
      var1: this.Servicedetails[0].var1,
      var2: this.Servicedetails[0].var2,
      var3: '',
      var1Value: this.Servicedetails[0].var1Value,
      var2Value: this.Servicedetails[0].var2Value,
      var3Value: '',
      PaytypeC: this.Servicedetails[0].PaytypeC,
      PaytypeW: this.Servicedetails[0].PaytypeW,
      PaytypeI: this.Servicedetails[0].PaytypeI,
      DepartmentS: this.Servicedetails[0].DepartmentS,
      DepartmentP: this.Servicedetails[0].DepartmentP,
      DepartmentQ: this.Servicedetails[0].DepartmentQ,
      DepartmentB: this.Servicedetails[0].DepartmentB,
      PolicyAccount: this.Servicedetails[0].PolicyAccount,
      excludeZeroHours: this.Servicedetails[0].zeroHours,
      PageNumber: 0,
      PageSize: '10000',
      // "startdealdate":this.Servicedetails[0].StartDate,
      // "enddealdate":this.Servicedetails[0].EndDate,
      // "var1": "Store_Name",
      // "var2": "ServiceAdvisor_Name",
      // "var3": "",
      // "var1Value":this.Servicedetails[0].var1Value,
      // "var2Value":this.Servicedetails[0].var2Value,
      // "var3Value":this.Servicedetails[0].var3Value
    };
    this.shared.api
      .postmethod(this.shared.common.routeEndpoint + 'GetServicesGrossSummaryDetailsV2', obj)
      .subscribe((res) => {
        if (res.status == 200) {

          this.shared.spinner.hide()
          if (res.response != undefined) {
            if (res.response.length > 0) {
              let localarray = res.response
              const workbook = this.shared.getWorkbook();
              const worksheet = workbook.addWorksheet('Service Gross Details');
              worksheet.views = [
                {
                  state: 'frozen',
                  ySplit: 9, // Number of rows to freeze (2 means the first two rows are frozen)
                  topLeftCell: 'A10', // Specify the cell to start freezing from (in this case, the third row)
                  showGridLines: false,
                },
              ];
              const DateToday = this.shared.datePipe.transform(new Date(), 'MM/dd/yyyy h:mm:ss a');

              const titleRow = worksheet.getCell("A2"); titleRow.value = 'Service Gross Details';
              titleRow.font = { name: 'Arial', family: 4, size: 15, bold: true };
              titleRow.alignment = { indent: 1, vertical: 'middle', horizontal: 'left' }



              const DateBlock = worksheet.getCell("L2"); DateBlock.value = DateToday;
              DateBlock.font = { name: 'Arial', family: 4, size: 10 };
              DateBlock.alignment = { vertical: 'middle', horizontal: 'center' }
              worksheet.addRow([''])
              const Store_Name = worksheet.addRow(['Store Name :']);
              Store_Name.getCell(1).font = { name: 'Arial', family: 4, size: 9, bold: true, };
              Store_Name.alignment = { indent: 1, vertical: 'middle', horizontal: 'left' }
              const StoreName = worksheet.getCell("B4"); StoreName.value = this.Servicedetails[0].var1Value;
              StoreName.font = { name: 'Arial', family: 4, size: 9 };
              StoreName.alignment = { indent: 1, vertical: 'middle', horizontal: 'left' }

              const DATE_EXTENSION = this.shared.datePipe.transform(
                new Date(),
                'MMddyyyy'
              );

              const StartDealDate = worksheet.addRow(['Start Date :']);
              const startdealdate = worksheet.getCell('B5');
              startdealdate.value = this.Servicedetails[0].StartDate;
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
              enddealdate.value = this.Servicedetails[0].EndDate;
              enddealdate.font = { name: 'Arial', family: 4, size: 9 };
              enddealdate.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
              EndDealDate.getCell(1).font = {
                name: 'Arial',
                family: 4,
                size: 9,
                bold: true,
              };

              const Var1Value = worksheet.addRow(['Advisor Name :']);
              Var1Value.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
              const var1value = worksheet.getCell('B7');
              var1value.value = this.Servicedetails[0].userName == '' ? '-' : this.Servicedetails[0].userName == null ? '-' : this.Servicedetails[0].userName;
              var1value.font = { name: 'Arial', family: 4, size: 9 };
              var1value.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
              Var1Value.getCell(1).font = {
                name: 'Arial',
                family: 4,
                size: 9,
                bold: true,
              };
              worksheet.addRow('');
              let Headings = [
                'Sl.no',
                'RO #',
                'Date',
                'Customer',
                'Vehicle',
                'VIN',
                'Op Code',
                'Op Code Desc',
                'Total Gross',
                'Total Labor',
                'Total Parts',
                'Total Tires',
                'Total Lube',
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
              for (const d of localarray) {
                count++
                d.closedate = this.shared.datePipe.transform(d.closedate, 'MM/dd/yyyy');
                const Data1 = worksheet.addRow([
                  count,
                  d.ronumber == '' ? '-' : d.ronumber == null ? '-' : d.ronumber,
                  d.closedate == '' ? '-' : d.closedate == null ? '-' : d.closedate,
                  d.CName == '' ? '-' : d.CName == null ? '-' : d.CName,
                  d.vehicle == '' ? '-' : d.vehicle == null ? '-' : d.vehicle,
                  d.vin == '' ? '-' : d.vin == null ? '-' : d.vin,
                  d.ASD_Opcode == '' ? '-' : d.ASD_Opcode == null ? '-' : d.ASD_Opcode,
                  d.ASD_opcodedescription == '' ? '-' : d.ASD_opcodedescription == null ? '-' : d.ASD_opcodedescription,

                  d.Totalgross == '' ? '-' : d.Totalgross == null ? '-' : parseFloat(d.Totalgross),
                  d.Totallabour == '' ? '-' : d.Totallabour == null ? '-' : parseFloat(d.Totallabour),
                  d.TotalParts == '' ? '-' : d.TotalParts == null ? '-' : parseFloat(d.TotalParts),
                  d.TotalTires == '' ? '-' : d.TotalTires == null ? '-' : parseFloat(d.TotalTires),
                  d.TotalBulkOil == '' ? '-' : d.TotalBulkOil == null ? '-' : parseFloat(d.TotalBulkOil),
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
                  // Apply default border and number format
                  cell.border = { right: { style: 'thin' } };
                  cell.numFmt = '$#,##0'; // Default format
                  if (number == 9 || number == 20 || number == 25 || number == 30) {
                    cell.numFmt = '$#,##0.00'; // Default format
                    cell.alignment = { indent: 1, vertical: 'middle', horizontal: 'center', };
                  }
                  if (number == 18 || number == 23 || number == 28 || number == 33) {
                    cell.numFmt = '0.00';
                    cell.alignment = { indent: 1, vertical: 'middle', horizontal: 'center', };
                  }
                  if (number == 17 || number == 22 || number == 27 || number == 32) {
                    cell.numFmt = '###0.00';
                    cell.alignment = { indent: 1, vertical: 'middle', horizontal: 'center', };
                  }
                  if (number > 6 && number < 27) {
                    cell.alignment = { indent: 1, vertical: 'middle', horizontal: 'center', };
                  }

                  else if (number == 1) {
                    cell.numFmt = '###0';
                    cell.alignment = { indent: 1, vertical: 'middle', horizontal: 'center', };
                  } else if (number == 2) {
                    cell.alignment = { indent: 1, vertical: 'middle', horizontal: 'center', };
                  } else if (number > 1 && number < 7) {
                    cell.alignment = { indent: 1, vertical: 'middle', horizontal: 'center', };
                  }



                  // Apply specific formatting for number 1
                  if (number === 1) {
                    cell.numFmt = '###0'; // No decimals
                    cell.alignment = {
                      indent: 1,
                      vertical: 'middle',
                      horizontal: 'center',
                    };
                  }

                  // Alignment for number 2
                  if (number === 2) {
                    cell.alignment = {
                      indent: 1,
                      vertical: 'middle',
                      horizontal: 'center',
                    };
                  }

                  // Alignment for numbers between 1 and 7 (excluding 1)
                  if (number > 1 && number < 7) {
                    cell.alignment = {
                      indent: 1,
                      vertical: 'middle',
                      horizontal: 'center',
                    };
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

              this.shared.exportToExcel(workbook, 'Service Gross Details' + DATE_EXTENSION);

            }
          }
        } else {
          this.shared.spinner.hide()
          // this.toast.show(res.error,'danger','Error')
        }
      })
  }
}