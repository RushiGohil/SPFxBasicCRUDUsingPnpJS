import * as React from 'react';
import styles from './SpFxBasicCrudUsingPnPjs.module.scss';
import { ISpFxBasicCrudUsingPnPjsProps } from './ISpFxBasicCrudUsingPnPjsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { getSP } from '../pnpjsConfig';

import { PrimaryButton, Stack, MessageBar, MessageBarType } from 'office-ui-fabric-react';
import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { IItemAddResult } from "@pnp/sp/items";

//create state
export interface ISpFxBasicCrudUsingPnPjsState {
  showmessageBar: boolean; //to show/hide message bar on success
  message: string; // what message to be displayed in message bar
  itemID: number; // current item ID after create new item is clicked
}

export default class SpFxBasicCrudUsingPnPjs extends React.Component<ISpFxBasicCrudUsingPnPjsProps, ISpFxBasicCrudUsingPnPjsState> {

  public _sp: SPFI;
  private listName = "DemoForSPFxBasicCRUD";

  constructor(props: ISpFxBasicCrudUsingPnPjsProps, state: ISpFxBasicCrudUsingPnPjsState) {
    super(props);
    this.state = { showmessageBar: false, message: "", itemID: 0 };
    this._sp = getSP();
  }

  public render(): React.ReactElement<ISpFxBasicCrudUsingPnPjsProps> {
    return (
      <div>
        <Stack horizontal tokens={{ childrenGap: 40 }}>
          <PrimaryButton text="Create New Item" onClick={() => this.createNewItem()} />
          <PrimaryButton text="Get Item" onClick={() => this.getItem()} />
          <PrimaryButton text="Update Item" onClick={() => this.updateItem()} />
          <PrimaryButton text="Delete Item" onClick={() => this.delteItem()} />
        </Stack>

        <br></br>
        <br></br>
        {
          this.state.showmessageBar &&
          <MessageBar onDismiss={() => this.setState({ showmessageBar: false })}
            dismissButtonAriaLabel="Close">
            {this.state.message}
          </MessageBar>
        }
      </div>
    );
  }

  // method to use pnp objects and create new item
  private async createNewItem() {
    const iar: IItemAddResult = await this._sp.web.lists.getByTitle(this.listName).items.add({
      Title: "Title " + new Date(),
      Description: "This is item created using PnP JS"
    });
    console.log(iar);
    this.setState({ showmessageBar: true, message: "Item Added Sucessfully", itemID: iar.data.Id });
  }

  // method to use pnp objects and get item by id, using item ID set from createNewItem method.
  private async getItem() {
    // get a specific item by id
    const item: any = await this._sp.web.lists.getByTitle(this.listName).items.getById(this.state.itemID)();
    console.log(item);
    this.setState({ showmessageBar: true, message: "Last Item Created Title:--> " + item.Title });
  }

  // method to use pnp object udpate item by id, using item id set from createNewItem method.
  private async updateItem() {
    let list = this._sp.web.lists.getByTitle(this.listName);
    const i = await list.items.getById(this.state.itemID).update({
      Title: "My Updated Title",
      Description: "Here is a updated description"
    });
    console.log(i);
    this.setState({ showmessageBar: true, message: "Item updated sucessfully" });
  }

  // method to use pnp object udpate item by id, using item id set from createNewItem method.
  private async delteItem() {
    let list = this._sp.web.lists.getByTitle(this.listName);
    var res = await list.items.getById(this.state.itemID).delete();
    console.log(res);
    this.setState({ showmessageBar: true, message: "Item deleted sucessfully" });
  }
}
