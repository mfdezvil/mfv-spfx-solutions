import { IMailItem } from "../models/IMailItem";
import { ISPGraphMails } from "../entities/ISPGraphMail";
import { Utils } from "../utils/Utils";
import * as strings from 'GraphConnectionTestWebPartStrings';

export class MailHelper {

    public static GetMailsFromResponse(response: ISPGraphMails, currentCulture:string): IMailItem[] {
        var emails: Array<IMailItem> = new Array<IMailItem>();
        response.value.map((item: any) => {
            let emailDate: string = (!Utils.IsNullOrEmpty(item.receivedDateTime)) ? new Date(item.receivedDateTime).toLocaleDateString(currentCulture) : "";
            emails.push({
                received: emailDate,
                from: item.from["emailAddress"].name + "("+item.from["emailAddress"].address+")",
                subject: item.subject,
                isRead: (item.isRead) ? strings.Yes : strings.No 
            });
        });
        return emails;
    }


}