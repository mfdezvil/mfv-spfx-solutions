import { ISPGraphTeamChannel } from "./ISPGraphTeamChannel";

export interface ISPGraphTeams {
    value: ISPGraphTeam[];
}

export interface ISPGraphTeam {
    id: string;
    displayName: string;
    description: string;
    channels: ISPGraphTeamChannel[];
}