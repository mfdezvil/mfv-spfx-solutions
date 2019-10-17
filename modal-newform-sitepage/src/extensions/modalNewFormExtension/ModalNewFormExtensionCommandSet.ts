import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';

import { setup as pnpSetup } from "@pnp/common";
import  ModalDialog from './components/ModalDialog';
import * as strings from 'ModalNewFormExtensionCommandSetStrings';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IModalNewFormExtensionCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'ModalNewFormExtensionCommandSet';

export default class ModalNewFormExtensionCommandSet extends BaseListViewCommandSet<IModalNewFormExtensionCommandSetProperties> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized ModalNewFormExtensionCommandSet');
    return super.onInit().then(_ => {

      pnpSetup({
        spfxContext: this.context
      });
    });
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'COMMAND_1':
        const dialog: ModalDialog = new ModalDialog();
        dialog.show();
        break;
      default:
        throw new Error('Unknown command');
    }
  }
}
