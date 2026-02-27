import { Component, ElementRef, Renderer2, ViewChild, signal } from '@angular/core';
import { SidebarService } from '../../Core/Providers/Shared/sidebar.service (1)';
import { environment } from '../../../environments/environment';
import { common } from '../../common';
import { Api } from '../../Core/Providers/Api/api';
import { Router } from '@angular/router';
import { NgFor, NgIf } from '@angular/common';
import { NgStyle, NgClass, NgForOf } from '@angular/common';
Api

import { HttpClient } from '@angular/common/http';

import { RivIcon } from '../..//Core/Providers/Shared/riv-icon/riv-icon';
@Component({
  selector: 'app-externalmenu',
  imports: [NgFor, NgIf, NgStyle, NgClass, NgForOf, NgIf],
  templateUrl: './externalmenu.html',
  styleUrl: './externalmenu.scss'
})
export class Externalmenu {
  submenu3: any;
  Report = false;
  submenu2 = false;
  @ViewChild('menu') menu: ElementRef | undefined;
  @ViewChild('sidebar') sidebar: ElementRef | undefined;
  uname: any = '';
  userObj: any = '';
  subid: any = 0;
  supersub: any = 0;
  activeid: any = 0;
  isSidebarCollapsed: any = false;

  // public userInfo: any;
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
  public userAuthData: any;
  constructor(
    public renderer: Renderer2,
    public apiService: Api,
    private router: Router,
    public comm: common,
    private sidebarService: SidebarService //private helpmodalservice: HelpmodalService
  ) {
    // Only global event listener in constructor
    this.renderer.listen('document', 'click', (event: Event) => this.handleClickOutside(event));
  }


  ngOnInit(): void {
    let cleanUrl = window.location.href;
    cleanUrl = cleanUrl.replace(/[?&]token=[^&#]*/g, '').replace(/[?&]$/, '');
    if (cleanUrl.includes('#')) {
      const [base, hash] = cleanUrl.split('#');
      let cleanHash = hash.replace(/[?&]token=[^&#]*/g, '').replace(/[?&]$/, '');
      cleanUrl = base + (cleanHash ? `#${cleanHash}` : '');
    }
    console.log('🧹 Cleaning token from URL:', cleanUrl);
    window.history.replaceState({}, document.title, cleanUrl);
    this.isSidebarCollapsed = !!localStorage.getItem('sidebarCollapsed');
    this.uname = localStorage.getItem('AO_uname') || '';
    const user = localStorage.getItem('userobj');
    this.userObj = user ? JSON.parse(user) : null;
    const storedUserInfo = localStorage.getItem('userInfo');
    this.userInfo = storedUserInfo ? JSON.parse(storedUserInfo) : null;
    this.gettoken()
    setTimeout(() => {
      this.initializeHeaderMenu();
    }, 0);
  }
  ngAfterViewInit(): void {
    // Ensure correct initial state after route changes
    this.isSidebarCollapsed = this.sidebarService.getCollapsed();
  }

  initializeHeaderMenu(): void {
    if (!this.userInfo || !this.userInfo?.user_Info?.roleid) {
      console.warn('UserInfo not yet available. Retrying getheadermenu...');
      setTimeout(() => this.initializeHeaderMenu(), 0);
      return;
    }

    this.apiService.changehead(true);

    this.getheadermenu();
    // this.getdata();
  }
  disabledParentIds: Set<number> = new Set();

  parentmenu: any;

  gettoken() {

    const body = {
      userid: this.userInfo.user_Info.ADuserid
    };

    this.apiService.postmethodtoken('fredbeans/AxelOneUserinfo', body)
      .subscribe((res: any) => {

        if (res.status === 200) {
          const api = res.response;
          this.apiService.setUserToken(api);

          const tokenData = this.userInfo;
          const finalUser = {
            userid: tokenData.user_Info.ADuserid,
            role: api.roleid || 1,
            session: {
              username: tokenData.fullName,
              email: tokenData.user_Info.email,
              role: tokenData.user_Info.roleid,
              oth_stores: tokenData.user_Info.oth_stores,
              payroll_info: api.payroll_info,
              iat: tokenData.user_Info.iat,
              exp: tokenData.user_Info.exp
            },
            store: api.userstores || tokenData.userstores,
            flag: "M",
            pin: "",
            other_stores: api.oth_stores || tokenData.oth_stores
          };

          console.log("FINAL USER:", finalUser);
          this.converttotoken(finalUser)
        }

      });
  }
  converttotoken(data: any) {
    const sessionStr = JSON.stringify(data.session);
    const encodedSession = btoa(sessionStr);
    data.session = encodedSession;
    const fullStr = JSON.stringify(data);
    const encodedFull = btoa(fullStr);
    localStorage.setItem("finalUser", encodedFull);
    console.log("DOUBLE ENCODED USER STORED:", encodedFull);
  }


  getdata(storedModules: any) {
    this.apiService.getMethod('cmsmodules/getAOparentmodule').subscribe((res: any) => {
      const responseArray = Array.isArray(res) ? res : Object.values(res);
      const filtered = responseArray
        .filter((item: any) => item.aopm_status === 'Y')
        .sort((a: any, b: any) => a.aopm_seq - b.aopm_seq);
      const parentMenuArray: any[] = filtered.map((item: any) => [{
        Mod_ID: item.aopm_id,
        mod_name: item.aopm_title,
        mod_prod_id: item.aopm_id,
        prod_title: item.aopm_title,
        mod_seq: item.aopm_seq
      }]);
      this.parentmenu = parentMenuArray.sort((a, b) => a[0].mod_seq - b[0].mod_seq);
      const storedTitles = storedModules
        .flat()
        .map((m: any) => m.prod_title?.toLowerCase().trim())
        .filter(Boolean);
      this.disabledParentIds.clear();
      this.parentmenu.forEach((arr: any[]) => {
        const item = arr[0];
        const id = item.mod_prod_id;
        if (!storedTitles.includes(item.prod_title?.toLowerCase().trim()) && id !== 1 && id !== 2) {
          this.disabledParentIds.add(id);
        }
      });
    });
  }

  selectedSubmenuId: any = 0;

  selectedModule: any;
  toggleReport(id: any): void {

    if (id == '') return;
    this.supersub = 0;

    // Toggle off if already selected
    if (this.subid === id) {
      this.subid = 0;
      this.prodid = 0;
      this.activeid = 0;
      this.selectedModule = null;
      this.selectedSubmenuId = 0; // reset highlight
      return;
    }

    this.subid = id;
    this.prodid = id;

    if (id !== 6) {
      this.selectedModule = this.comm.modules.find((item: any) => item[0].mod_prod_id === id);

      if (this.selectedModule) {
        this.supersub = this.selectedModule[0].Mod_ID;
        this.activeid = this.selectedModule[0].Mod_ID;

        // ✅ Auto-select first submenu (other than header)
        const firstSub = this.selectedModule.find(
          (sub: any) => sub.Mod_ID && sub.mod_name !== this.selectedModule[0].prod_title
        );
        if (firstSub) {
          this.selectedSubmenuId = firstSub.Mod_ID;
        } else {
          this.selectedSubmenuId = 0; // fallback
        }
      }

      console.log('select module', this.selectedModule);
    }
  }

  prodid: any = 0;
  togglesubmenu2(event: any, id: any, prodid: any) {
    event.stopPropagation();
    this.prodid = prodid;
    this.supersub = id;
    this.submenu2 = true;
    this.activeid = id;
    this.selectedSubmenuId = id;
    var elems = document.querySelectorAll('li');
    [].forEach.call(elems, function (el: any) {
      el.className = el.className.replace(/\bactive\b/, '');
    });

    this.renderer.addClass(event.target, 'active');
  }

  close() {
    var elems = document.querySelectorAll('li');

    this.submenu2 = false;
    this.supersub = 0;
  }
  isCollapsed = false;

  toggleSidebar() {

    this.isSidebarCollapsed = !this.isSidebarCollapsed;
    this.sidebarService.toggle();
    this.closemenu();
    this.closesub();
    this.toggleMainMenuclose();
    this.toggleSubMenuclose();
  }

  handleClickOutside(event: Event) {
    const targetElement = event.target as HTMLElement;

    const clickedInsideMenu = this.menu?.nativeElement.contains(targetElement);
    const clickedInsideSidebar = this.sidebar?.nativeElement.contains(targetElement);

    if (!clickedInsideMenu && !clickedInsideSidebar) {
      this.closemenu(); // 👈 close both menus
      this.toggleMainMenuclose();
      this.toggleSubMenuclose();
    }
  }
  // gotoLink(url: any) {
  //   const moduleTitle = this.selectedModule?.[0]?.prod_title?.toLowerCase().trim();
  //   console.log('module', moduleTitle);
  //   const firstProdTitle: string = this.selectedModule?.[0]?.prod_title?.toLowerCase().trim() || '';
  //   console.log(firstProdTitle); // "reports"

  //   this.closemenu();
  //   this.closesub();

  //   if (!url) return;

  //   try {
  //     const uToken = localStorage.getItem('Login_token') || '';
  //     const currentOrigin = window.location.origin;
  //     const currentHost = window.location.host;
  //     const isLocal = currentHost.includes('localhost') || currentHost.includes('127.0.0.1');

  //     const [baseUrl, hashPart] = url.split('#');
  //     const hasHash = !!hashPart;
  //     const target =
  //       baseUrl.startsWith('http://') || baseUrl.startsWith('https://')
  //         ? new URL(baseUrl)
  //         : new URL(baseUrl, currentOrigin);

  //     target.searchParams.set('token', uToken);

  //     let finalUrl = target.toString();
  //     if (hasHash) finalUrl += `#${hashPart}`;

  //     const targetHost = target.host;

  //     // ✅ CASE 1: Reports module → Always open in new tab, even with hash
  //     if (moduleTitle === 'reports') {
  //       console.log('Opening Reports module in new tab (even with hash):', finalUrl);
  //       window.open(finalUrl, '_blank');
  //       return;
  //     }

  //     // ✅ CASE 2: Localhost navigation (other modules)
  //     if (isLocal) {
  //       if (moduleTitle === 'reports') {
  //         console.log('Opening Reports module in new tab (even with hash):', finalUrl);
  //         window.open(finalUrl, '_blank');
  //         return;
  //       }
  //       console.log('moduletitle', moduleTitle);

  //       if (moduleTitle !== 'reports') {
  //         console.log('Running on localhost');
  //         this.router.navigateByUrl(target.pathname + target.search).catch(() => {
  //           console.warn(`⚠️ Localhost: Route not found, directly loading ${finalUrl}`);
  //           window.location.href = finalUrl;
  //         });
  //         return;
  //       }
  //     }

  //     // ✅ CASE 3: External host (different domain)
  //     if (targetHost !== currentHost) {
  //       window.open(finalUrl, '_blank');
  //       return;
  //     }

  //     // ✅ CASE 4: Normal internal navigation
  //     if (hasHash) {
  //       window.open(finalUrl, '_blank');
  //     } else {
  //       this.router.navigateByUrl(target.pathname + target.search).catch((err) => {
  //         console.warn(`⚠️ Route not found, redirecting directly to: ${finalUrl}`, err);
  //         window.location.href = finalUrl;
  //       });
  //     }
  //   } catch (error) {
  //     console.error('❌ Invalid URL:', error, 'for input:', url);
  //   }
  // }
  gotoLink(url: any) {
    const moduleTitle = this.selectedModule?.[0]?.prod_title?.toLowerCase().trim();
    console.log('module', moduleTitle);
    const firstProdTitle: string = this.selectedModule?.[0]?.prod_title?.toLowerCase().trim() || '';
    console.log(firstProdTitle);

    this.closemenu();
    this.closesub();

    if (!url) return;

    try {
      const uToken = localStorage.getItem('finalUser') || '';
      const currentOrigin = window.location.origin;
      const currentHost = window.location.host;
      const isLocal = currentHost.includes('localhost') || currentHost.includes('127.0.0.1');

      const [baseUrl, hashPart] = url.split('#');
      const hasHash = !!hashPart;
      const target =
        baseUrl.startsWith('http://') || baseUrl.startsWith('https://')
          ? new URL(baseUrl)
          : new URL(baseUrl, currentOrigin);

      target.searchParams.set('token', uToken);

      let finalUrl = target.toString();
      if (hasHash) finalUrl += `#${hashPart}`;

      const targetHost = target.host;

      // ✅ CASE 1: Reports module → Always open in new tab, even with hash
      if (moduleTitle === 'reports') {
        console.log('Opening Reports module in new tab (even with hash):', finalUrl);
        if (window.location.href.includes('localhost')) {
          let endpoint = finalUrl.replace(this.comm.menuUrl, '');
          this.router.navigateByUrl(endpoint);
        } else {
          window.location.href = finalUrl;
        }
        return;
      }

      // ✅ CASE 2: Localhost navigation (other modules)
      if (isLocal) {
        if (moduleTitle === 'reports') {
          console.log('Opening Reports module in new tab (even with hash):', finalUrl);
          window.open(finalUrl, '_blank');
          return;
        }
        console.log('moduletitle', moduleTitle);

        if (moduleTitle !== 'reports') {
          console.log('Running on localhost');
          this.router.navigateByUrl(target.pathname + target.search).catch(() => {
            console.warn(`⚠️ Localhost: Route not found, directly loading ${finalUrl}`);
            window.location.href = finalUrl;
          });
          return;
        }
      }

      // ✅ CASE 3: External host (different domain)
      if (targetHost !== currentHost) {
        window.open(finalUrl, '_blank');
        return;
      }

      // ✅ CASE 4: Normal internal navigation
      if (hasHash) {
        window.open(finalUrl, '_blank');
      } else {
        this.router.navigateByUrl(target.pathname + target.search).catch((err) => {
          console.warn(`⚠️ Route not found, redirecting directly to: ${finalUrl}`, err);
          window.location.href = finalUrl;
        });
      }
    } catch (error) {
      console.error('❌ Invalid URL:', error, 'for input:', url);
    }
  }


  getheadermenu() {
    this.comm.modules = [];

    this.apiService.headermenu(`${this.userInfo.user_Info.roleid}`).subscribe(
      (data: any) => {
        console.log('Get Header Menu  : ', data);
        let obj =
          data != '' && data != undefined
            ? this.groupArrayOfObjects(data.response, 'mod_prod_id')
            : [];
        if (obj != '') {
          let map = Object.values(obj).forEach((e: any) => {
            if (e.length > 0) {
              let mapped = e.map((ev: any) => {
                let ob = {
                  Mod_ID: ev.Mod_ID,
                  mod_name: ev.mod_name,
                  mod_prod_id: ev.mod_prod_id,
                  prod_title: ev.prod_title,
                  xmlData: this.extractData(ev.xmlData),
                };
                return ob;
              });
              this.comm.modules.push(mapped);
            }
          });
        }
        console.log(this.comm.modules);
        this.getdata(this.comm.modules);
        this.getplaybooksdata();
        localStorage.setItem('modulename', JSON.stringify(this.comm.modules));
      },
      (err: any) => {
        this.comm.modules = [];
      }
    );
  }

  groupArrayOfObjects(list: any, key: any) {
    return list.reduce(function (rv: any, x: any) {
      (rv[x[key]] = rv[x[key]] || []).push(x);
      return rv;
    }, {});
  }

  extractData(str: any) {
    str = str.replace('"[', '[');
    str = str.replace(']"', ']');
    let Json = JSON.parse(str);
    //console.log(Json);
    return Json;
  }

  closesub() {
    this.submenu2 = false;
  }

  closemenu() {
    this.subid = 0;
    this.supersub = 0;
    this.selectedModule = '';
  }

  isEqualIgnoreCase(a: string | any, b: string | any): boolean {
    if (a === null || a === undefined || b === null || b === undefined) return false;
    return a.toString().trim().toLowerCase() === b.toString().trim().toLowerCase();
  }
  hasValidSubmenuForItem(modules: any[], supersub: number): boolean {
    for (let item of modules) {
      for (let sub of item) {
        if (sub.Mod_ID === supersub && !this.isEqualIgnoreCase(sub.mod_name, item[0].prod_title)) {
          return true;
        }
      }
    }
    return false;
  }

  handleClick(item: any[]) {
    const first = item[0];
    const second = item[1];
    this.closemenu();
    this.closesub();
    this.toggleMainMenuclose();
    this.toggleSubMenuclose();
    // Case 1: Dashboard
    if (first?.prod_title?.toLowerCase() === 'dashboard') {
      window.location.href = ' https://fb.axelone.app/dashboard';
      return;
    }

    // Case 2: Ask Axel
    if (first?.prod_title?.toLowerCase() === 'ask axel') {
      // if you want ask axel
      // this.navigateToUrl(second?.mod_url);
      return;
    }

    // if (first?.prod_title?.toLowerCase() === 'customers') {
    //   const ut = JSON.stringify(this.userInfo);
    //   const uToken = btoa(ut);
    //   window.location.href = `https://fb.axelone.app/cc?token=${uToken}`;
    //   return;
    // }

    if (first?.mod_filename && first.mod_filename.toLowerCase().startsWith('https')) {
      this.router.navigate(['/']);
      return;
    }

    // Case 4: Normal behavior
    if (first?.mod_prod_id === 2) {
      this.router.navigate(['/']);
    } else {
      this.toggleReport(first.mod_prod_id);
    }
  }

  navigateToUrl(query: string) {
    const baseUrl = localStorage.getItem('webToken') || '';
    const encodedQuery = encodeURIComponent(query);
    const fullUrl = `${baseUrl}&role=${this.userInfo?.user_Info?.roleid}&group=${this.userInfo.group}`;
    window.open(fullUrl, '_blank');
  }

  // playbooksmenu:any;
  // playbookselectedModule:boolean=false;
  // getplaybooksdata(){
  //   let obj = {
  //     "roleid":this.userInfo?.user_Info?.roleid
  //   }
  //   this.apiService.postmethodHRDB('playbooks/GetPlaybooksMenuV2',obj).subscribe((res: any) => {
  //     console.log(res);
  //     if(res.status == 200){
  //       let playbooksmenu = res.response;
  //        let mapped = playbooksmenu.map((ev: any) => {
  //               let ob = {
  //                 FREQUENCY: ev.FREQUENCY,
  //                 FreqName: ev.FreqName,
  //                 Module_List: this.extractData(ev.Module_List),
  //               };
  //               return ob;
  //             });
  //       this.playbooksmenu= mapped;
  //       console.log(this.playbooksmenu)
  //     }
  //   });
  // }

  playbookmainmenu: any
  playbooksmenu: any;
  playbookselectedModule: boolean = false;
  getplaybooksdata() {
    let obj = {
      "roleid": this.userInfo?.user_Info?.roleid
    }
    this.apiService.postmethodHRDB('playbooks/GetPlaybooksMenuV2', obj).subscribe((res: any) => {
      console.log(res);
      if (res.status == 200) {
        let playbooksmenu = res.response;
        let finobj: any = {};
        finobj = this.groupDataByFrequency(playbooksmenu);
        this.playbookmainmenu = Object.keys(finobj).map(key => {
          return {
            label: this.getLabel(key),
            items: finobj[key],
            isSelected: false
          };
        });
      }
    });
  }

  getLabel(key: string): string {
    const map: any = {
      D: 'Daily',
      W: 'Weekly',
      M: 'Monthly'
    };
    return map[key] || key;
  }


  toggleMainMenu(menu: any) {
    this.playbooksmenu = (menu.items || []).map((x: any) => ({
      name: x.NAME,
      Module_List: x.MOdules,
      isSelected: false,
      count: x.MOdules?.length || 0
    }));
    if (Array.isArray(this.playbookmainmenu)) {
      this.playbookmainmenu.forEach((m: any) => {
        if (m !== menu) m.isSelected = false;
      });
    }

    menu.isSelected = !menu.isSelected;
    this.togglepbSubMenu(this.playbooksmenu[0]);
  }

  togglepbSubMenu(sub: any) {
    this.playbooksmenu.forEach((s: any) => {
      if (s !== sub) s.isSelected = false;
    });

    sub.isSelected = !sub.isSelected;
  }

  toggleMainMenuclose(): void {
    this.playbookmainmenu.forEach((i: any) => {
      i.isSelected = false;
    });
  }

  toggleSubMenuclose(): void {
    this.playbooksmenu?.forEach((i: any) => i.isSelected = false);
  }

  getselectedTitle() {
    const selected = this.playbookmainmenu?.find((m: any) => m.isSelected);
    return selected ? selected.label : '';
  }

  isAnyMenuSelected() {
    return this.playbookmainmenu?.some((m: any) => m.isSelected);
  }

  isAnySubMenuSelected() {
    return this.playbooksmenu?.some((m: any) => m.isSelected);
  }

  gotoLinkPlaybooks(url: any) {
    this.toggleMainMenuclose();
    this.toggleSubMenuclose();
    if (!url) return;

    try {
      const uToken = localStorage.getItem('finalUser') || '';
      const currentOrigin = window.location.origin;

      const [baseUrl, hashPart] = url.split('#');
      const hasHash = !!hashPart;
      const target =
        baseUrl.startsWith('http://') || baseUrl.startsWith('https://')
          ? new URL(baseUrl)
          : new URL(baseUrl, currentOrigin);

      target.searchParams.set('token', uToken);

      let finalUrl = target.toString();
      if (hasHash) finalUrl += `#${hashPart}`;

      window.open(finalUrl, '_blank');
    }
    catch (error) {
      console.error('❌ Invalid URL:', error, 'for input:', url);
    }
  }

  groupedData: any = {};

  groupDataByFrequency(data: any) {
    const result = data.reduce((acc: any, item: any) => {
      const key = item.FREQUENCY2;
      if (!acc[key]) {
        acc[key] = [];
      }
      if (item.MOdules) {
        try {
          item.MOdules = JSON.parse(item.MOdules);
        } catch (e) {
          console.error("JSON Parse Error:", e);
          item.MOdules = [];
        }
      }
      acc[key].push(item);
      return acc;
    }, {});
    this.groupedData = result;
    return result;
  }
}
