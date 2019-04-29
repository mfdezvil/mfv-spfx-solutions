import { MSGraphClient, MSGraphClientFactory } from '@microsoft/sp-http';
import { MSGraphClientVersions } from '../utils/MSGraphClientUtils';
import { ISPGraphUser } from '../entities/ISPGraphUser';
import { Logger } from '../utils/Logger';
import { IGraphDataReader } from '../managers/IGraphDataReader';
import { ISPGraphMails } from '../entities/ISPGraphMail';
import { ISPGraphTeams } from '../entities/ISPGraphTeam';
import { ISPGraphTeamChannels } from '../entities/ISPGraphTeamChannel';

export class GraphRequestService implements IGraphDataReader { 

    private msGraphClientFactory: MSGraphClientFactory = null;
    private graphClient: MSGraphClient = null;

    constructor (msGraphClientFactory: MSGraphClientFactory) {
      this.msGraphClientFactory = msGraphClientFactory;
      this.Init();
    }

    private async Init() {
      this.graphClient = await this.msGraphClientFactory.getClient();
    }

    public async GetCurrentUserData():Promise<ISPGraphUser> {
      try{
        var p = new Promise<ISPGraphUser>(async (resolve, reject) => { 
          var selectedProperties:string = "displayName,mail,userPrincipleName";

          let callback = (error, response) => {
              if (error) {
                  Logger.WriteError("GetCurrentUserData", error.message);
                  reject(error);
              } else {
                  resolve(response);
              }
          };

          await this.graphClient
            .api("me")
            .version(MSGraphClientVersions.V1)
            .select(selectedProperties)
            .get(callback);
        });
        return p;
      } 
      catch (error) {
        Logger.WriteError("GetCurrentUserData", error);
        return null;
      }
    }

    public async GetCurrentUserMails():Promise<ISPGraphMails> {
      try{ 
        var p = new Promise<ISPGraphMails>(async (resolve, reject) => { 
          var selectedProperties:string = "receivedDateTime,from,subject,isRead";
          var filters: string = "isDraft eq false";
          var orderby: string = "receivedDateTime DESC";

          let callback = (error, response) => {
              if (error) {
                  Logger.WriteError("GetCurrentUserMails", error.message);
                  reject(error);
              } else {
                  resolve(response);
              }
          };

          await this.graphClient
            .api("me/messages")
            .version(MSGraphClientVersions.V1)
            .select(selectedProperties)
            .filter(filters)
            .orderby(orderby)
            .get(callback);
        });
        return p;
      } 
      catch (error) {
        Logger.WriteError("GetCurrentUserMails", error);
        return null;
      }
    }

    public GetCurrentUserEvents():Promise<any> {
      try{
        var p = new Promise<any>(async (resolve, reject) => { 
          var selectedProperties:string = "organizer,subject,start,end,location";
          var orderby:string = "createdDateTime DESC";

          let callback = (error, response) => {
              if (error) {
                  Logger.WriteError("GetCurrentUserEvents", error.message);
                  reject(error);
              } else {
                  resolve(response);
              }
          };
          await this.graphClient
          .api("me/events")
          .version(MSGraphClientVersions.V1)
          .select(selectedProperties)
          .orderby(orderby)
          .get(callback);
        });
        return p;
      } 
      catch (error) {
        Logger.WriteError("GetCurrentUserEvents", error);
        return null;
      }
    }
    
    public GetCurrentUserTeams():Promise<ISPGraphTeams> {
      try{
        var p = new Promise<ISPGraphTeams>(async (resolve, reject) => { 
          var selectedProperties:string = "id,displayName,description";

          let callback = (error, response) => {
              if (error) {
                  Logger.WriteError("GetCurrentUserTeams", error.message);
                  reject(error);
              } else {
                  resolve(response);
              }
          };
          await this.graphClient
          .api("/me/joinedTeams")
          .version(MSGraphClientVersions.V1)
          .select(selectedProperties)
          .get(callback);
        });
        return p;
      } 
      catch (error) {
        Logger.WriteError("GetCurrentUserTeams", error);
        return null;
      }
    }

    public GetChannelsByTeam(teamID: string): Promise<ISPGraphTeamChannels> {
      try{
        var p = new Promise<ISPGraphTeamChannels>(async (resolve, reject) => { 
          var selectedProperties:string = "id,displayName,description,email,webUrl";

          let callback = (error, response) => {
              if (error) {
                  Logger.WriteError("GetChannelsByTeam", error.message);
                  reject(error);
              } else {
                  resolve(response);
              }
          };
          await this.graphClient
          .api("/teams/"+teamID+"/channels")
          .version(MSGraphClientVersions.V1)
          .select(selectedProperties)
          .get(callback);
        });
        return p;
      } 
      catch (error) {
        Logger.WriteError("GetChannelsByTeam", error);
        return null;
      }
    }

}