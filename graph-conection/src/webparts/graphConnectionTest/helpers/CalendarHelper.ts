import { IEventItem } from "../models/IEventItem";
import { Utils } from "../utils/Utils";

export class CalendarHelper {

    public static GetEventsFromResponse(response: any, currentCulture: string): IEventItem[] {
        var events: Array<IEventItem> = new Array<IEventItem>();
        response.value.map((item: any) => {
            let start: string = (!Utils.IsNullOrEmpty(item.start.dateTime)) ? new Date(item.start.dateTime).toLocaleString(currentCulture) : "";
            let end: string = (!Utils.IsNullOrEmpty(item.end.dateTime)) ? new Date(item.end.dateTime).toLocaleString(currentCulture) : "";
            events.push({
                organizer: item.organizer["emailAddress"].name + " ("+item.organizer["emailAddress"].address+")",
                subject: item.subject,
                start: start,
                end: end,
                location: item.location.address.street + ", "+  item.location.address.city
            });
        });
        return events;
    }


}