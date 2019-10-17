import { Web, Items } from "@pnp/sp";
import { IBaseListItem, ISPListItemOrder } from "../entities/SharepointListsServiceEntities";
import { ISPListsServiceRequest } from "../entities/SharepointListsServiceRequests";
import { Utils } from "../utils/Utils";
import { StringHelper } from "../helpers/StringHelper";

export class SharepointListsService {

    private currentWeb: Web = null;

    constructor (webURL:string) {
        this.currentWeb = new Web(webURL);
    }

    public async ReadItems<T extends IBaseListItem>(request: ISPListsServiceRequest): Promise<Array<T>> {
        let queryWeb: Web = this.GetRequestWeb(request.ListUrl);
        let query: Items;
        if(!Utils.IsNullOrEmpty(request.ListGuid)) {
            query = queryWeb.lists.getById(request.ListGuid).items;
        }
        else {
            query= queryWeb.lists.getByTitle(request.ListName).items;
        }

        let fieldsToSelect: string = request.GetSelect();
        query = query.select(fieldsToSelect);

        let expand: string = request.GetExpand();
        if(!Utils.IsNullOrEmpty(expand)) {
            query= query.expand(expand);
        }

        let filters: string = request.GetFilters();
        if(!Utils.IsNullOrEmpty(filters)) {
            query = query.filter(filters);
        }
        
        let order: ISPListItemOrder = request.GetOrder();
        if(!Utils.IsNullOrEmpty(order)) {
            query= query.orderBy(order.ColumnName, order.OrderType);
        }

        let limit: number = request.GetLimit();
        if(!Utils.IsNullOrEmpty(limit))
            query= query.top(limit);

        return await query.get();
    }

    private GetRequestWeb(requestListUrl: string): Web {
        if(!Utils.IsNullOrEmpty(requestListUrl)) {
            let listUrl= StringHelper.SplitString(requestListUrl, "/Lists/");
            return new Web(listUrl);
        }
        else {
            return this.currentWeb;
        }
    }
    

}