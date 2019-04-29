import { ISPGraphUser } from "../entities/ISPGraphUser";
import { ISPGraphMails } from "../entities/ISPGraphMail";
import { ISPGraphTeams } from "../entities/ISPGraphTeam";
import { ISPGraphTeamChannels } from "../entities/ISPGraphTeamChannel";

export interface IGraphDataReader {
    GetCurrentUserData(): Promise<ISPGraphUser>;
    GetCurrentUserMails(): Promise<ISPGraphMails>;
    GetCurrentUserTeams(): Promise<ISPGraphTeams>;
    GetCurrentUserEvents(): Promise<any>;
    GetChannelsByTeam(teamID: string): Promise<ISPGraphTeamChannels>;
}