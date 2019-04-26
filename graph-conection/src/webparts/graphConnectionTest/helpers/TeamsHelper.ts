import { ITeamItem } from "../models/ITeamItem";

export class TeamsHelper {
   
    public static GetTeamsFromResponse(response: any): ITeamItem[] {
        var teams: Array<ITeamItem> = new Array<ITeamItem>();
        response.value.map((item: any) => {
            teams.push({
                displayName: item.displayName,
                description: item.description,
                id: item.id
            });
        });
        return teams;
    }

}