import { ITeamItem } from "../models/ITeamItem";
import { IChannelItem } from "../models/IChannelItem";
import { ISPGraphTeamChannel } from "../entities/ISPGraphTeamChannel";
import { ISPGraphTeam } from "../entities/ISPGraphTeam";

export class TeamsHelper {
   
    public static GetTeamsFromResponse(response: any): ITeamItem[] {
        var teams: Array<ITeamItem> = new Array<ITeamItem>();
        response.value.map((item: ISPGraphTeam) => {
            let teamChannels: IChannelItem[] = [];
            item.channels.map((channel: ISPGraphTeamChannel) => {
                teamChannels.push({
                    id: channel.id,
                    displayName: channel.displayName,
                    description: channel.description,
                    email: channel.email,
                    webUrl: channel.webUrl
                });
            });
            teams.push({
                displayName: item.displayName,
                description: item.description,
                id: item.id,
                channels: teamChannels
            });
        });
        return teams;
    }

}