import { ITestListItem } from "../models/ITestListItem";

export interface ITestWebpartState {
    text: string;
    items: Array<ITestListItem>;
}