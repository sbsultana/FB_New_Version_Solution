import { Component, EventEmitter, HostListener, Input, Output, SimpleChanges } from '@angular/core';
import { common } from '../../common';
import { SharedModule } from '../../Core/Providers/Shared/shared.module';


@Component({
  selector: 'app-stores',
  imports: [SharedModule],
  standalone: true,
  templateUrl: './stores.html',
  styleUrl: './stores.scss'
})
export class Stores {
  @Input() storesFilterData: any = {}
  @Output() StoresData = new EventEmitter();
  @Output() storeLable = new EventEmitter();


  storeIds: any = []
  stores: any = []
  groupName: any = ''
  groupsArray: any = [];
  storename: any = '';
  groupId: any;
  type: any = '';
  others: any = '';
  storedisplayname: any = ''
  storecount: any = null


  constructor(private comm: common) { }

  ngOnChanges(changes: SimpleChanges) {
    console.log(changes);

    if (changes['storesFilterData'] && changes['storesFilterData'].currentValue) {
      this.type = changes['storesFilterData'].currentValue.type;
      this.others = changes['storesFilterData'].currentValue.others;
      this.groupsArray = changes['storesFilterData'].currentValue.groupsArray;
      this.stores = changes['storesFilterData'].currentValue.storesArray;

      this.storeIds = changes['storesFilterData'].currentValue.storeids;
      this.groupId = changes['storesFilterData'].currentValue.groupId;
      this.groupName = changes['storesFilterData'].currentValue.groupName;
      this.storename = changes['storesFilterData'].currentValue.storename;
      this.storecount = changes['storesFilterData'].currentValue.storecount;
      this.storedisplayname = changes['storesFilterData'].currentValue.storedisplayname;
      console.log(this.stores,changes['storesFilterData'].currentValue.storesArray);
      
      if (this.storesFilterData.storesArray && this.storesFilterData.storesArray.length > 0) {
        this.getGroupBaseStores(this.groupId.toString())
      }

    }
  }
  ngOnInit() {
    // this.getStoresandGroupsValues()

  }
  individualgroups(e: any) {
    this.groupId = []
    this.groupId.push(e.sg_id);
    this.groupName = this.groupsArray.filter((val: any) => val.sg_id == this.groupId)[0].sg_name;
    this.getGroupBaseStores(this.groupId.toString())
  }
  getGroupBaseStores(id: any) {
    this.stores = this.comm.groupsandstores.filter((v: any) => v.sg_id == id)[0]?.Stores;
    console.log(this.storeIds, 'Store IDs');
    let singleid=this.storeIds
    this.storeIds = []
    this.type == 'S' ? this.storeIds.push(parseInt(singleid)) : this.storeIds = this.stores.map(function (a: any) {
      return a.ID;
    });
    console.log(this.storeIds,'After condition');
    
    if (this.storeIds && this.storeIds.length == 1) {
      this.storename = this.stores.filter((val: any) => val.ID == this.storeIds.toString())[0].storename;
    } else {
      this.storename = ''
    }
    this.groupName = this.groupsArray.filter((val: any) => val.sg_id == this.groupId)[0].sg_name;
    this.setValues()
  }
  individualStores(e: any) {
    const index = this.storeIds.findIndex((i: any) => i == e.ID);
    if (index >= 0) {
      this.storeIds.splice(index, 1);
    } else {
      this.type == 'S' ? this.storeIds = [] : ''
      this.type == 'S' ? this.storeIds.push(e.ID) : this.storeIds.push(e.ID);
    }
    if (this.storeIds.length == 1) {
      this.storename = this.stores.filter((val: any) => val.ID == this.storeIds.toString())[0].storename
    } else {
      this.storename = ''
    }
    this.setValues()

  }
  allstores(state: any) {
    if (state == 'N') {
      this.storeIds = [];
    } else if (state == 'Y') {
      this.storeIds = this.stores.map(function (a: any) {
        return a.ID;
      });
      this.groupName = this.groupsArray.filter((val: any) => val.sg_id == this.groupId)[0].sg_name;
    }
    this.setValues()
  }

  setValues() {

    if (this.storeIds.length == 1) {
      this.storecount = null;
      this.storedisplayname = this.storename
    }
    else if (this.storeIds.length == this.stores.length) {
      this.storecount = null;
      this.storedisplayname = this.groupName;
    }
    else if (this.storeIds.length > 1) {
      this.storecount = this.storeIds.length;
      this.storedisplayname = 'Selected'
    }
    else {
      this.storedisplayname = 'Select'
    }
    this.storesFilterData.groupsArray = this.groupsArray;
    this.storesFilterData.storesArray = this.stores;

    this.storesFilterData.groupId = this.groupId;
    this.storesFilterData.storeids = this.storeIds;
    this.storesFilterData.storename = this.storename;
    this.storesFilterData.groupName = this.groupName;
    this.storesFilterData.storecount = this.storecount;
    this.storesFilterData.storedisplayname = this.storedisplayname;

    this.StoresData.emit(this.storesFilterData);

  }


}