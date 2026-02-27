import { Component, ElementRef, EventEmitter, inject, Output, PLATFORM_ID, Renderer2, ViewChild } from '@angular/core';
import { Sharedservice } from '../../Core/Providers/Shared/sharedservice';
import { Setdates } from '../../Core/Providers/SetDates/setdates';
import { ActivatedRoute, Router, NavigationStart, RouterModule, } from '@angular/router';
import { decode } from 'js-base64';
import { isPlatformBrowser } from '@angular/common';
import { SharedModule } from '../../Core/Providers/Shared/shared.module';
import { UserNetworkService } from '../../Core/Providers/Shared/global.services';


import { environment, UserInfo } from '../../../environments/environment';
import { NgFor, NgIf } from '@angular/common';
import { NgStyle, NgClass, NgForOf } from '@angular/common';
Api
ToastService
import { HttpClient } from '@angular/common/http';
NotificationService
import { RivIcon } from '../..//Core/Providers/Shared/riv-icon/riv-icon';
import { Api } from '../../Core/Providers/Api/api';
import { ToastService } from '../../Core/Providers/Shared/toast.service';
import { NotificationService } from '../../Core/Providers/Shared/notification.service';

import { AuthServices } from '../../Core/Providers/Shared/auth.service 1';
import { common } from '../../common';

declare var bootstrap: any;
@Component({
  selector: 'app-header',
  imports: [SharedModule, RivIcon, NgIf, RouterModule],
  standalone: true,
  templateUrl: './header.html',
  styleUrl: './header.scss'
})


export class Header {
  public componentData: any;
  public loading = false;
  private toast = inject(ToastService);
  // public userInfo : any;
  private router = inject(Router);
  dealerGroups: string[] = ['Cavender', 'Fredbeans', 'Price'];
  KeyWords = ['Prasad', 'Chavali', 'Pchavali', 'Greedily', 'Greedily Routine4'];
  selectedGroup: string | null = null;
  groupsandstoresdata: any = []
  stores: any = [];
  userInfo: any =
  {
    "group": "FREDBEANS",
    "fullName": "Prasad Chavali",
    "GU_URL": "https://fb.axelone.app/",
    "xtract_url": "https://fbxtract..axelone.app/",
    "user_Info": {
        "userid": 1,
        "firstname": "Prasad",
        "lastname": "Chavali",
        "email": "pchavali-axelone@fredbeans.com",
        "phone": "+19000000004",
        "roleid": 100,
        "status": "Y",
        "profileImg": "88a8b6b1a8270a1eda07c76cab267cb9prasad.png",
        "title": "Super Admin",
        "moduleids": "1,2,53,4,5,8,16,18,19,22,23,24,26,27,28,29,30,35,40,48,62,66,67,70,71,72,73,74,75,76,77,78,79,80,88,89,90,91,92,93,94,95,96,97,99,100,102,104,7,11,12,68,14,15,25,81,83,84,85,87,98,101,103,41,36,63,69,54",
        "preferences": 1,
        "ADuserid": "DEALERS\\PChavali-AxelOne",
        "storeids": "71,8,7,4,35,1,32,40,50,25,18,31,3,70,72,2,17,41,42,51,12,73,9,15,5,14,30,11,53,55,54",
        "userstores": "71,8,7,4,35,1,32,40,50,25,18,31,3,70,72,2,17,41,42,51,12,73,9,15,5,14,30,11,53,55,54",
        "ustores": "71,8,7,4,35,1,32,40,50,25,18,31,3,70,72,2,17,41,42,51,12,73,9,15,5,14,30,11,53,55,54",
        "oth_stores": "",
        "uid": 1,
        "Xtract": 1,
        "Touch": 1,
        "Xperience": 1,
        "Xchange": 3,
        "Xiom": 1,
        "Tracs": 1,
        "iat": 1767102955,
        "exp": 1798638955
    },
    "site": "prod"
}
  ceoLogs: any = [
    {
      username: 'dealers\\pchavali-axelone',
      group: 'FREDBEANS',
      stgroup: 'FREDBEANS',
      url: 'https://fredbeans.axelautomotive.com/',
      pwd: 'AxelOne1!',
      logo: 'FB_Logo.png'
    },
    {
      username: 'prasad.chavali@axelautomotive.com',
      group: 'WESTERN',
      stgroup: 'WESTERN',
      url: 'https://westernauto.axelone.app/',
      pwd: 'Cav@2024$$',
      logo: 'WesternLogo.svg'
    },
    {
      username: 'prasad.chavali@axelautomotive.com',
      group: 'TPG',
      stgroup: 'TPG',
      url: 'https://tropical.axelone.app',
      pwd: 'tcp@2025$$',
      logo: 'Tropical_Logo.png'
    },
    {
      username: 'MOSSY\\AxelOne',
      group: 'MOSSY',
      stgroup: 'MOSSY',
      url: 'https://mossy.axelone.app',
      pwd: 'Welcome2Mossy!',
      logo: 'Mossy_Logo.png'
    },
    // {
    //   username : 'prasad.chavali@axelautomotive.com',
    //   group : 'Dream',
    //   stgroup : 'DMG1',
    //   url : 'https://dream.axelautomotive.com/',
    //   pwd : 'Dmg@2024$$',
    //   logo : 'Dream_Logo.png'
    // },
    // {
    //   username : 'prasad.chavali@axelautomotive.com',
    //   group : 'Ancira',
    //   stgroup : 'ANCIRA',
    //   url : 'https://ancira.axelautomotive.com/',
    //   pwd : 'Anc@2024$$',
    //   logo : 'Ancira_logo.png'
    // },
    // {
    //   username : 'PRICESIMMS\\axel1',
    //   group : 'Price',
    //   stgroup : 'Price',
    //   url : 'https://price.axelautomotive.com/',
    //   pwd : 'Rug-Dubbed-Pavement2',
    //   logo : 'Price_Logo.png'
    // },
    // {
    //   username : 'RigoG@toyotacedarpark.com',
    //   group : 'Idea',
    //   stgroup : 'ID',
    //   url : 'https://idea.axelautomotive.com/',
    //   pwd : 'axelone'
    // },
    // {
    //   username : 'MWAXEL3@motorwerks.com',
    //   group : 'Murgado',
    //   stgroup : 'MG',
    //   url : 'https://murgado.axelautomotive.com/',
    //   pwd : 'Murgad0@2024'
    // },
  ]

  hasUnread = false;
  matched: boolean = false;
  constructor(public shared: Sharedservice, public AuthServices: AuthServices, public userNetworkInfo: UserNetworkService, public apiCall: Api, private http: HttpClient, public notificationService: NotificationService, public comm: common, public apiService: Api,) {

    //this.userNetworkInfo.networkStatus = false;
    // const storedUser = localStorage.getItem('userInfo');
    // const storedUser = this.userInfo
    // if (userInfo) {
    const userInfo = this.userInfo;
    localStorage.setItem('userInfo', JSON.stringify(this.userInfo))

    console.log('I Am Header : ', userInfo);

    this.userInfo = userInfo;
    //console.log(userInfo.fullName); // "Veera Babu C"

    if (this.userInfo?.user_Info) {
      const fullName = this.userInfo?.fullName?.toLowerCase() || '';
      const firstName = this.userInfo.user_Info?.firstname?.toLowerCase() || '';
      const lastName = this.userInfo.user_Info?.lastname?.toLowerCase() || '';
      this.matched = this.KeyWords.some(keyword =>
        fullName.includes(keyword.toLowerCase()) ||
        firstName.includes(keyword.toLowerCase()) ||
        lastName.includes(keyword.toLowerCase())
      );
    }

    const tropicalname = this.userInfo.fullName?.toLowerCase()

    this.matched = this.KeyWords.some(keyword =>
      tropicalname.includes(keyword.toLowerCase()));




    // }else{
    //     var userInfo = this.userNetworkInfo.getUser();
    //     this.userInfo = userInfo;
    //     localStorage.setItem('userInfo', JSON.stringify(userInfo));

    //       const fullName = this.userInfo?.fullName?.toLowerCase() || '';
    //       const firstName = this.userInfo?.user_Info?.firstname?.toLowerCase() || '';
    //       const lastName = this.userInfo?.user_Info?.lastname?.toLowerCase() || '';


    //       this.matched = this.KeyWords.some(keyword =>
    //         fullName.includes(keyword.toLowerCase()) ||
    //         firstName.includes(keyword.toLowerCase()) ||
    //         lastName.includes(keyword.toLowerCase())
    //       );
    // }
    this.apiCall.GetHeaderData().subscribe((res: any) => {
      this.componentData = res.obj;
       this.shared.common.pageName = this.componentData.title
    }
    );
    console.log('Header Component Data : ', this.componentData.title);
    this.shared.common.pageName = this.componentData.title

  }

  ngOnInit(): void {
    const storedUser = localStorage.getItem('userInfo');
    console.log(storedUser, '....................');

    if (storedUser) {

      this.getNotificationCount();
      this.getGruopsandStores();
    }
  }

  onDealerSelect(dealer: any, index: number, event: Event): void {
    this.selectedGroup = dealer;

    this.loading = true;

    if (dealer.stgroup == 'CAV1' || dealer.stgroup == 'WESTERN' || dealer.stgroup == 'DMG1' || dealer.stgroup == 'MOSSY' || dealer.stgroup == 'FREDBEANS') {

      const postObj = {
        'username': dealer.username,
        'password': btoa(dealer.pwd),
        'group': dealer.stgroup
      };

      this.apiCall.postmethod('login/login2', postObj).subscribe({
        next: (res: any) => {

          if (res.status === 200) {

            const tkn = (dealer.stgroup == 'WESTERN' ? res.response.jwttoken : res.response)
            const userPayload = this.getJwtPayload(tkn);  // Safe extraction

            console.log('User info Header : ', userPayload);

            const userInfo = {
              group: res.group,
              user_aou_AD_userid: dealer.username,
              fullName: userPayload.firstname + ' ' + userPayload.lastname, // Use actual value if available
              GU_URL: res?.groupInfo?.GU_URL || '',
              xtract_url: res?.groupInfo?.XTRACT_URL || '',
              user_Info: userPayload,
              site: environment.site
            };

            //console.log('Before Local Storage : ', userInfo);

            this.userNetworkInfo.setUser(userInfo);
            this.userInfo = userInfo

            if (localStorage.getItem('userInfo')) {
              localStorage.removeItem('userInfo');
              localStorage.setItem('userInfo', JSON.stringify(userInfo));
            } else {
              localStorage.setItem('userInfo', JSON.stringify(userInfo));
            }

            if (userInfo.group == 'MOSSY') {
              var token = btoa(JSON.stringify(userInfo));

              console.log('User Info :', userInfo);
              if (!this.userNetworkInfo.networkStatus) {
                //alert(1);
                window.location.reload();
              } else {
                // return
                //window.location.href = `http://localhost:4201?token=${token}`;
                window.location.href = environment.mossyAppUrl + `?token=${token}`;
              }

            }
            else if (userInfo.group == 'FREDBEANS') {
              var token = btoa(JSON.stringify(userInfo));

              console.log('User Info :', userInfo);
              if (!this.userNetworkInfo.networkStatus) {
                //alert(1);
                window.location.reload();
              } else {
                // return
                window.location.href = environment.fbAppUrl + `?token=${token}`;
              }
            }
            else {
              window.location.reload();
            }


          } else {
            // this.toast.show('Invalid username or password', 'danger');
            this.toast.show('Invalid username or password', 'warning', 'Invalid');
          }

          //this.loading = false; // ✅ Always reset loading after `next`
        },
        error: (err: { error: { status: number; }; }) => {
          //console.error('Login error:', err);

          if (err.error.status && err.error.status == 401) {
            this.toast.show('Please try entering it again', 'warning', 'Passwword is Invalid');
          } else if (err.error.status && err.error.status == 404) {
            this.toast.show("We couldn't find an account with the eamil you entered", 'danger', 'Account not found');
          } else if (err.error.status && err.error.status == 503) {
            this.toast.show("Authentication service unavailable, Please try after some time!", 'warning', 'Service Unavaiable');
          }
          //this.toast.show('Something went wrong. Please try again.', 'danger');

          //this.loading = false; // ✅ Also reset loading on error
          //this.cdr.detectChanges();
        }
      });
    } else if (dealer.stgroup == 'TPG') {


      const postObj = {
        'username': dealer.username,
        'password': btoa(dealer.pwd),
        'group': dealer.stgroup
      };

      this.apiCall.getMethod(`login/status/a/${dealer.stgroup}/${dealer.username}`).subscribe({
        next: (data: any) => {

          console.log('Tpg User Login  : ', data);

          if (data.status == 200) {
            const tkn = data.response
            const userPayload = this.getJwtPayload(tkn);

            console.log('✅ API response 2:', userPayload);


            const userInfo = {
              group: dealer.stgroup,
              user_aou_AD_userid: userPayload?.user_Info?.ADuserid,
              fullName: userPayload?.firstname + ' ' + userPayload?.lastname, // Use actual value if available
              GU_URL: data?.groupInfo?.GU_URL || '',
              xtract_url: data?.groupInfo?.XTRACT_URL || '',
              user_Info: userPayload,
              xpertResp: '',
              site: environment.site
            };


            console.log('User info Obj : ', userInfo);

            this.userNetworkInfo.setUser(userInfo);
            this.userInfo = userInfo;
            if (localStorage.getItem('userInfo')) {
              localStorage.removeItem('userInfo');
              localStorage.setItem('userInfo', JSON.stringify(userInfo));

              window.location.reload();
            } else {
              localStorage.setItem('userInfo', JSON.stringify(userInfo));
              window.location.reload();
            }

            var token = btoa(JSON.stringify(userInfo));

            console.log('User Info :', userInfo);
            if (!this.userNetworkInfo.networkStatus) {
              //alert(1);
              window.location.reload();
            } else {
              // return
              //window.location.href = `http://localhost:4201?token=${token}`;
              window.location.href = environment.tropicalUrl + `?token=${token}`;
            }

            //this.understood();

          }


        }
      });
      // this.redirectToMicrosoftTropical('axel-testuser-1@tropicalchevrolet.com');

      // window.location.reload();
    }
    // else if (dealer.stgroup == 'MOSSY'){
    //     const userInfo = {
    //       group: dealer.stgroup,
    //       user_aou_AD_userid: dealer.username,
    //       fullName: dealer.username,
    //       GU_URL: dealer?.GU_URL || '',
    //       xtract_url: dealer?.XTRACT_URL || '',
    //       user_Info: ''
    //     };

    //     console.log('User Info :', userInfo)
    //     this.userNetworkInfo.setUser(userInfo);
    //     localStorage.setItem('userInfo', JSON.stringify(userInfo));

    //     const userinfos : any = this.userInfo;
    //     var token = btoa(JSON.stringify(userinfos));

    //    // window.location.href = `https://mossy.axelone.app?token=${token}`;
    //     // window.location.replace(`http://localhost:4321?token=${token}`);
    //     return;
    // }

    else {
      const postObj = {
        'username': dealer.username,
        'password': btoa(dealer.pwd),
      };


      this.apiCall.postmethod('login', postObj).subscribe({
        next: (res: any) => {
          //this.loading = false;
          //this.block = 'insight';
          //this.cdr.detectChanges();
          //console.log('API Response:', res);

          if (res.status === 200) {

            //this.toast.show('Login successful', 'success');

            //this.toast.show('Login Success.!', 'success', 'Success');



            const tkn = (dealer.stgroup == 'WESTERN' ? res.response.jwttoken : res.response)
            const userPayload = this.getJwtPayload(tkn);  // Safe extraction


            const userInfo = {
              group: res.group,
              user_aou_AD_userid: dealer.username,
              fullName: userPayload.firstname + ' ' + userPayload.lastname, // Use actual value if available
              GU_URL: res?.groupInfo?.GU_URL || '',
              xtract_url: res?.groupInfo?.XTRACT_URL || '',
              user_Info: userPayload,
              site: environment.site
            };

            //console.log('Before Local Storage : ', userInfo);

            this.userNetworkInfo.setUser(userInfo);
            if (localStorage.getItem('userInfo')) {
              localStorage.removeItem('userInfo');
              localStorage.setItem('userInfo', JSON.stringify(userInfo));
            } else {
              localStorage.setItem('userInfo', JSON.stringify(userInfo));
            }

            window.location.reload();
            // this.understood();

          } else {
            // this.toast.show('Invalid username or password', 'danger');
            this.toast.show('Invalid username or password', 'warning', 'Invalid');
          }

          // this.loading = false; // ✅ Always reset loading after `next`
        },
        error: (err: { error: { status: number; }; }) => {
          //console.error('Login error:', err);

          if (err.error.status && err.error.status == 401) {
            this.toast.show('Please try entering it again', 'warning', 'Passwword is Invalid');
          } else if (err.error.status && err.error.status == 404) {
            this.toast.show("We couldn't find an account with the eamil you entered", 'danger', 'Account not found');
          } else if (err.error.status && err.error.status == 503) {
            this.toast.show("Authentication service unavailable, Please try after some time!", 'warning', 'Service Unavaiable');
          }
          //this.toast.show('Something went wrong. Please try again.', 'danger');

          //this.loading = false; // ✅ Also reset loading on error
          //this.cdr.detectChanges();
        }

      });
    }
    // setTimeout(async () => {
    //   window.location.reload();
    // }, 2000);
  }

  redirectToMicrosoftTropical(userEmail: any): Promise<void> {

    return this.AuthServices.loginWithUsernameHint(userEmail); // Make sure 'initialize' is a public method in AuthServices
  }

  fetchDashboardData(dealer: string) {
    // Modify this method as per your backend API logic
    const obj = {
      dealerGroup: dealer,
      // other params
    };


  }

  dashbardRoute() {
    this.router.navigate(['/dashboard']);
  }

  logoutSession() {
    localStorage.removeItem('userInfo');
    this.router.navigate(['/']);
  }

  getJwtPayload(token: string): any {
    const parts = token.split('.');
    if (parts.length !== 3) {
      throw new Error('Invalid JWT format');
    }
    const payload = this.decodeBase64Url(parts[1]);
    return JSON.parse(payload);
  }

  decodeBase64Url(base64Url: string): string {
    // Replace - with + and _ with / to make it valid Base64
    const base64 = base64Url.replace(/-/g, '+').replace(/_/g, '/');
    // Add padding if needed
    const padded = base64.padEnd(base64.length + (4 - (base64.length % 4)) % 4, '=');
    return atob(padded);
  }

  getNotificationCount() {
    this.notificationService.unreadCount$.subscribe((count) => {
      console.log('Count of notificaiton  : ', count);
      this.hasUnread = count > 0 ? true : false;
    });
  }


  onNotificationIconClick() {
    // alert('Hello !');

    const offcanvasEl = document.getElementById('notif');
    if (offcanvasEl) {
      const offcanvas = new bootstrap.Offcanvas(offcanvasEl);
      offcanvas.show();
    }
    this.notificationService.triggerFetch(); // Notify dashboard
  }


  onExcelClick() {

    const Data = {
      title: this.componentData.title,
      state: true,
    };
    this.apiCall.setExportToExcelAllReports({ obj: Data });

  }
  getGruopsandStores() {
    this.shared.common.completeUserDetails = this.userInfo
    // alert('HI')
    const obj = {
      "userid": JSON.parse(localStorage.getItem('userInfo')!).user_Info.userid,
      // "userid": 44,
    }
    this.shared.api.postmethod(this.shared.common.routeEndpoint + 'GetStoresList', obj)
      .subscribe((res) => {
        if (res.status == 200) {
          this.shared.common.allstores = res.response
          let data = res.response
          // .filter((val: any) => val.sg_id != 7)
          this.groupsandstoresdata = data.reduce((r: any, { sg_name, sg_id }: any) => {
            if (!r.some((o: any) => o.sg_name == sg_name)) {
              r.push({
                sg_name,
                sg_id,
                Stores: data.filter((v: any) => v.sg_name == sg_name),
              });
            }
            return r;
          }, []);
          console.log(this.groupsandstoresdata);
          // .filter((val:any)=>{ val.sg_name != 'Other Stores' })
          this.shared.common.groupsandstores = this.groupsandstoresdata.filter((val: any) => val.sg_name != 'OtherStores')
          this.shared.common.OtherStoresData = this.groupsandstoresdata.filter((val: any) => val.sg_name == 'OtherStores')
          // .filter((val:any)=>{ val.sg_name == 'Other Stores' })
          const obj = {
            storesData: this.shared.common.groupsandstores,
            otherStoresData: this.shared.common.OtherStoresData
          };
          this.shared.api.setStores({ obj: obj });
        }
      })
  }
}


