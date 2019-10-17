import * as React from 'react';
import {IModalDialogContentProps} from './IModalDialogContentProps';
import {IModalDialogContentState} from './IModalDialogContentState';
import {
    Dialog,
    DialogFooter,
    DefaultButton,
    PrimaryButton,
    ChoiceGroup,
    DialogType,
    TextField,
    Dropdown,
    IDropdownOption,
    Spinner,
    SpinnerSize,
    Label,
    DatePicker, 
    DayOfWeek
} from 'office-ui-fabric-react';
//import * as strings from 'ModalNewFormExtensionCommandSetStrings';
import {NewsTypeOptions, DateStringsOptions} from '../utils/Constants';
import {PageHelper} from '../helpers/PageHelper';
import { INewsPage } from '../models/INewsPage';

export default class ModalDialogContent extends React.Component<IModalDialogContentProps, IModalDialogContentState> {

    private newsTypeOptions: IDropdownOption[] = [
        { key: NewsTypeOptions.CategoriaA, text: NewsTypeOptions.CategoriaA },
        { key: NewsTypeOptions.CategoriaB, text: NewsTypeOptions.CategoriaB }
      ];

    constructor(props: IModalDialogContentProps) { 
        super(props);
        let newsData: INewsPage = {
            newPageName: "",
            newsPageType: null,
            newsExpirationDate: null
        };
        this.state = {
            hideDialog: false,
            optionSelected: 'A',
            newsPageData: newsData,
            isLoading: false,
            creationDone: false
        };
    }

    public render() {

        return (
            <div>
                <Dialog
                    hidden={this.state.hideDialog}
                    onDismiss={this.closeDialog}
                    dialogContentProps={{
                        type: DialogType.normal,
                        title: "Create news page"
                    }}
                    modalProps={{
                        isBlocking: true,
                        containerClassName: 'ms-dialogMainOverride'
                    }}
                >
                    {!this.state.creationDone && (
                        <div>
                        <TextField label="Page name" value={this.state.newsPageData.newPageName} onChanged={this.onChangeNewPageName} placeholder="Insert page name"
                            readOnly={this.state.isLoading}
                        />
                        <Dropdown
                                label="News Page Type"
                                selectedKey={this.state.newsPageData.newsPageType}
                                placeholder="Select news type"
                                onChanged={this.onChangeNewsType}
                                options={this.newsTypeOptions}
                                disabled={this.state.isLoading}
                        />
                        <DatePicker
                            label="Expiration Date"
                            firstDayOfWeek={DayOfWeek.Monday}
                            strings={DateStringsOptions.DayPickerStrings}
                            placeholder="Select a date..."
                            ariaLabel="Select a date"
                            value={this.state.newsPageData.newsExpirationDate}
                            onSelectDate={this.onChangeNewsExpirationDate}
                        />
                        
                        <div>
                            <DialogFooter>
                            <PrimaryButton onClick={this.executeAction} text="Go!" />
                            <DefaultButton onClick={this.closeDialog} text="Cancel" />
                            </DialogFooter>
                        </div>
                        </div>

                    )}
                    {this.state.isLoading && (<Spinner size={SpinnerSize.large} label="Please, wait.." ariaLive="assertive" hidden={!this.state.isLoading} />
                    )}
                    {this.state.creationDone && (
                        <div>
                        <Label>Done!</Label>
                        <DialogFooter>
                            <PrimaryButton onClick={this.closeDialog} text="Close" />
                        </DialogFooter>
                        </div>
                    )}
                </Dialog>
            </div>
        );
    }

    private executeAction = (): void => {
        this.setState({ isLoading: true });
        var result: Promise<string> = PageHelper.createCustomPage(this.state.newsPageData);
        result.then(ss => {
          console.log(ss);
          this.setState({ isLoading: false, creationDone: true });
        });
    }

    private onChangeNewsType = (option: IDropdownOption): void => {
        let newsData: INewsPage = this.state.newsPageData;
        newsData.newsPageType = option.key.toString();
        this.setState({ newsPageData: newsData });
    }

    private onChangeNewPageName = (tmpPageName: any): void => {
        let newsData: INewsPage = this.state.newsPageData;
        newsData.newPageName = tmpPageName;
        this.setState({ newsPageData: newsData });
    }

    private onChangeNewsExpirationDate = (date: Date | null | undefined): void => {
        let newsData: INewsPage = this.state.newsPageData;
        newsData.newsExpirationDate = date;
        this.setState({ newsPageData: newsData });
    }

    private showDialog = (): void => {
        this.setState({ hideDialog: false });
    }
    
    private closeDialog = (): void => {
        this.props.close();
        this.setState({ hideDialog: true });
    }
}