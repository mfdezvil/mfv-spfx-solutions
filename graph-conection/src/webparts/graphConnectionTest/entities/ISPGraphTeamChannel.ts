export interface ISPGraphTeamChannels {
    value: ISPGraphTeamChannel[];
}

export interface ISPGraphTeamChannel {
    id: string;
    displayName: string;
    description: string;
    email: string;
    webUrl: string;
}