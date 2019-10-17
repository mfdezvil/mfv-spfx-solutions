import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import ModalDialogContent from './ModalDialogContent';

export default class ModalDialog extends BaseDialog {

    constructor() {
        super();
    }

    public render(): void {
    
        ReactDOM.render(<ModalDialogContent 
          close={ this.close.bind(this) }
    
        />, this.domElement);
    
    }

    public getConfig(): IDialogConfiguration {
        return {
          isBlocking: false
        };
    }

    public close(): Promise<void> {
        return super.close();
    }
    
    protected onAfterClose(): void {
        super.onAfterClose();
        super.close();
        // Clean up the element for the next dialog
        ReactDOM.unmountComponentAtNode(this.domElement);
    }
}