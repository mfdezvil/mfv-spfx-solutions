import { INewsPage } from "../models/INewsPage";

export interface IModalDialogContentState {
    hideDialog: boolean;
    optionSelected: string;
    newsPageData: INewsPage;
    isLoading: boolean;
    creationDone: boolean;
}