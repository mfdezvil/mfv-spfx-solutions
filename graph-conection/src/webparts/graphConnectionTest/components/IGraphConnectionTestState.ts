import { IUserItem } from "../models/IUserItem";
import { IMailItem } from "../models/IMailItem";
import { IEventItem } from "../models/IEventItem";
import { ITeamItem } from "../models/ITeamItem";


export interface IGraphConnectionTestState {
    currentUserData: IUserItem;
    userMails: IMailItem[];
    userEvents: IEventItem[];
    userTeams: ITeamItem[];
    isLoading: boolean;
    errorMessage: string;
}