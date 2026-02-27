import { Component, OnInit, Input, Renderer2, Output, EventEmitter, } from '@angular/core';
import { NgbActiveModal, NgbModal } from '@ng-bootstrap/ng-bootstrap';
import {Router} from '@angular/router'
import { Sharedservice } from '../../Core/Providers/Shared/sharedservice';
import { SharedModule } from '../../Core/Providers/Shared/shared.module';



@Component({
  selector: 'app-notes',
  imports: [SharedModule],
  standalone:true,
  templateUrl: './notes.html',
  styleUrl: './notes.scss'
})
export class Notes {
  @Input() notesData: any;
  @Output() onClose = new EventEmitter();
  // @Output() SavedNotesData = new EventEmitter();


  userid: any = ''
  spinnerLoader: boolean = false;
  notes: any = '';
  menu: any = []
  curl: any = '';
  renderer: any;

  constructor( public shared:Sharedservice,private router: Router,  ) {
    // this.renderer.listen('window', 'click', (e: Event) => {
    //   const TagName = e.target as HTMLButtonElement;
    // });
    this.curl = router.url;
  }

  ngOnInit(): void {
    this.userid = JSON.parse(localStorage.getItem('UserDetails')!);
    let menuData = this.shared.common.menuData.flat();
    this.menu = menuData.filter((v: any) => v.mod_filename.indexOf(this.curl) > 0)[0].mod_id
    //console.log(this.userid, localStorage.getItem('UserDetails'), this.notesData);
  }


  close(val: any) {
    // this.ngbmodel.dismissAll();
    this.onClose.emit('C');

  }

  savenotes() {
    if (this.notes == '') {
      alert('Please enter notes');
    } else {
      let obj = {}
      if (this.notesData.apiRoute == 'AddGeneralNotes') {
        obj = {
          "GN_AS_ID": this.notesData.store,
          "GN_Title1": this.notesData.title1,
          "GN_Title2": this.notesData.title2,
          "GN_Text": this.notes,
          "GN_NS_ID": '',
          "GN_Module_ID": 49,
          "GN_Active": "Y",
          "GN_UserId": JSON.parse(localStorage.getItem('userInfo')!)?.user_Info?.userid,
        };
      }else if (this.notesData.apiRoute == 'AddNotesAction'){
        obj = {
          "AS_ID": this.notesData.store,
          "Title": this.notesData.mainkey,
          "Module":49,
          "Notes": this.notes,
          'UserID': JSON.parse(localStorage.getItem('userInfo')!)?.user_Info?.userid,
        };
      }
       
      this.spinnerLoader = true
      this.shared.api.postmethod(this.shared.common.routeEndpoint+this.notesData.apiRoute, obj).subscribe(
        (res) => {
          if (res.status == 200) {
           alert('Notes inserted successfully')
            // let data={notes:this.notes+'- ' +JSON.parse(localStorage.getItem('UserDetails')!).UserName +' - ('+ this.datepipe.transform(new Date(),'MM/dd/yy') + ')' }
            // this.SavedNotesData.emit(data);
            this.onClose.emit('S');
            // this.ngbmodel.dismissAll();

          } else {
            alert('Invalid Details');
          }
        },
        (error) => { }
      );
    }
  }







}
