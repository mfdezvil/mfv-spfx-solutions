import { ISPGraphUser } from "../entities/ISPGraphUser";
import { ISPGraphMails } from "../entities/ISPGraphMail";
import { ISPGraphTeams } from "../entities/ISPGraphTeam";

export interface IGraphDataReader {
    GetCurrentUserData(): Promise<ISPGraphUser>;
    GetCurrentUserMails(): Promise<ISPGraphMails>;
    GetCurrentUserTeams(): Promise<ISPGraphTeams>;
    GetCurrentUserEvents(): Promise<ISPGraphTeams>;
}