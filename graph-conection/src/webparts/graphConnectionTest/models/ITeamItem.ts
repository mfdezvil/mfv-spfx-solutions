import { IChannelItem } from "./IChannelItem";

export interface ITeamItem {
    id: string;
    displayName: string;
    description: string;
    channels: IChannelItem[];
}