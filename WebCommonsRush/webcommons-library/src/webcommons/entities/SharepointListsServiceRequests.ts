import { ISPColumn, LookupSPColumn, SPBasicColumn, PeopleSPColumn, ISPListItemLimit, ISPListItemOrder, ISPListItemFilter, SPListLookupItemFilter, SPFilterType, SPFilterOperator } from "./SharepointListsServiceEntities";
import { StringHelper } from "../helpers/StringHelper";
import { Utils } from "../utils/Utils";

export interface ISPListsServiceRequest {
    ListName: string;
    ListUrl: string;
    ListGuid: string;
    GetSelect: () => string;
    GetExpand: () => string;
    GetFilters: () => string;
    GetOrder: () => ISPListItemOrder;
    GetLimit: () => number;
    
}

export class SPListRequest implements ISPListsServiceRequest  {
    public ListGuid: string;
    public ListName: string;
    public ListUrl: string;

    /**
    * Crea una request para la rest api de listas de SP
    * @param listName Nombre de la lista
    * @param listGuid Guid de la lista - Alternativo al nombre de la lista
    * @param listUrl Indica la ruta de la lista si está en otro sitio
    */
    constructor(listName?: string, listGuid?: string, listUrl?:string) {
        this.ListName = listName;
        this.ListGuid = listGuid;
        this.ListUrl = listUrl;
    }

    public GetSelect(): string {
        return "";
    }

    public GetExpand(): string {
        return "";
    }
    public GetFilters(): string {
        return "";
    }
    public GetOrder(): ISPListItemOrder {
        return null;
    }
    public GetLimit(): number {
        return null;
    }
}

export class SPListRequestItems extends SPListRequest {

    private expandStr: string = "";
    private isExpandCalculated: boolean = false;
    
    public FieldsToSelect: Array<ISPColumn>;
    public Limit: ISPListItemLimit;
    public Order: ISPListItemOrder;
    public Filters: Array<ISPListItemFilter>;
    
    public GetSelect(): string {
        var selectStr: string = "";
        if (this.FieldsToSelect && this.FieldsToSelect.length > 0) {
            for (var i in this.FieldsToSelect) {
                var field: ISPColumn = this.FieldsToSelect[i];
                if (field instanceof LookupSPColumn) {
                    field.ProjectedColumns.forEach((columnName: string) => {
                        selectStr += "," + (field as LookupSPColumn).LookupColumnName + "/" + columnName;
                    });
                    if(this.expandStr.indexOf(field.LookupColumnName) == -1)
                        this.expandStr += "," + field.LookupColumnName;
                }
                else if (field instanceof PeopleSPColumn) {
                    field.ProjectedColumns.forEach((columnName: string) => {
                        selectStr += "," + (field as PeopleSPColumn).LookupColumnName + "/" + columnName;
                    });
                    this.expandStr += "," + field.LookupColumnName + field.ExpandColumnSuffix;
                }
                else if (field instanceof SPBasicColumn) {
                    selectStr += "," + field.Name;
                }
                else selectStr += "," + field.toString();
            }
            selectStr = StringHelper.TrimChar(selectStr,",");
            this.expandStr = StringHelper.TrimChar(this.expandStr, ",");
        }
        this.isExpandCalculated = true;
        
        return selectStr;
    }

    public GetExpand(): string {
        if(!this.isExpandCalculated)
            this.GetSelect();
        return this.expandStr;
    }

    public GetFilters():string {
        var filterStr: string = null;
        if (this.Filters && this.Filters.length > 0) {
            this.Filters.forEach((filter, index) => {
                var property = filter.ColumnName;
                if (filter instanceof SPListLookupItemFilter) {
                    property += "/" + filter.LookupColumnName;
                }
                var newFilter = "(" + property + " " + SPFilterType[filter.FilterType] + " '" + encodeURI(filter.Value) + "')";
                if (index == 0) filterStr = newFilter;
                else filterStr += " " + SPFilterOperator[filter.FilterOperator] + " " + newFilter;
            });
        }
        return filterStr;
    }

    public GetOrder(): ISPListItemOrder {
        return (!Utils.IsNullOrEmpty(this.Order)) ? this.Order : null;
    }


    public GetLimit():number {
        return (!Utils.IsNullOrEmpty(this.Limit)) ? this.Limit.LimitNumber : null;
    }
    
}