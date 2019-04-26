export interface ISPGraphMails {
    value: ISPGraphMail[];
}

export interface ISPGraphMail {
    id: string;
    receivedDateTime: string;
    subject: string;
    from: ISPGraphMailUser;
    isRead: boolean;
}

export interface ISPGraphMailUser {
    emailAddress: ISPGraphMailUserInfo;
}

export interface ISPGraphMailUserInfo {
    address: string;
    name: string;
}