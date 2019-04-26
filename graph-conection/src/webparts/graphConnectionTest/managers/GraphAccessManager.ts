import {
    Environment,
    EnvironmentType
} from '@microsoft/sp-core-library';
import { MSGraphClientFactory } from '@microsoft/sp-http';
import { IGraphDataReader } from './IGraphDataReader';
import { GraphMock } from '../services/mocks/GraphMock';
import { GraphRequestService } from '../services/GraphRequestService';
import { IUserItem } from '../models/IUserItem';
import { CurrentUserDataHelper } from '../helpers/CurrentUserDataHelper';
import { IMailItem } from '../models/IMailItem';
import { MailHelper } from '../helpers/MailHelper';
import { ITeamItem } from '../models/ITeamItem';
import { TeamsHelper } from '../helpers/TeamsHelper';
import { IEventItem } from '../models/IEventItem';
import { CalendarHelper } from '../helpers/CalendarHelper';

export class GraphAccessManager {
    private graphDataReader: IGraphDataReader;
    private currentCulture: string;

    constructor(msGraphClientFactory: MSGraphClientFactory, currentCulture: string) {
        this.currentCulture= currentCulture;
        if (Environment.type === EnvironmentType.Local) {
            this.graphDataReader = new GraphMock();
        }
        else if (Environment.type == EnvironmentType.SharePoint || 
                  Environment.type == EnvironmentType.ClassicSharePoint) {
            this.graphDataReader = new GraphRequestService(msGraphClientFactory);
        }
    }

    public async GetUserData(): Promise<IUserItem> {
        
        var graphUser = await this.graphDataReader.GetCurrentUserData();
        
        return new Promise<IUserItem>((resolve) => {
            if(graphUser) {
                resolve(CurrentUserDataHelper.GetUserFromResponse(graphUser));
            }
            else resolve(null);
        });
    }

    private GetUserDataCallback(graphUser: IUserItem):IUserItem {
        if(graphUser) {
            return CurrentUserDataHelper.GetUserFromResponse(graphUser);
        }
        return null;
    }

    public async GetUserMails(): Promise<IMailItem[]> {
        
        var graphMails = await this.graphDataReader.GetCurrentUserMails();
        
        return new Promise<IMailItem[]>((resolve) => {
            if(graphMails) {
                resolve(MailHelper.GetMailsFromResponse(graphMails, this.currentCulture));
            }
            else resolve(null);
        });
    }

    public async GetUserEvents(): Promise<IEventItem[]> {
        
        var graphEvents = await this.graphDataReader.GetCurrentUserEvents();
        
        return new Promise<IEventItem[]>((resolve) => {
            if(graphEvents) {
                resolve(CalendarHelper.GetEventsFromResponse(graphEvents, this.currentCulture));
            }
            else resolve(null);
        });
    }

    public async GetUserTeams(): Promise<ITeamItem[]> {
        
        var graphTeams = await this.graphDataReader.GetCurrentUserTeams();
        
        return new Promise<ITeamItem[]>((resolve) => {
            if(graphTeams) {
                resolve(TeamsHelper.GetTeamsFromResponse(graphTeams));
            }
            else resolve(null);
        });
    }

}
