/* Entidades de respuesta de Sharepoint */

export interface IBaseListItem {
    ID?: number;
    Title?: string;
    Name?: string;
    Author?: ISharepointUser;
    BaseListItemNewProp?: string;
    Created?: string;
    Modified?: string;
    ModifiedBy?: ISharepointUser;
    OData__ModerationStatus?: number;
}

export interface ISharepointUser {
    Id?: number;
    Title?: string;
}

/* Entidades para el Select */

export interface ISPColumn { }

export class SPBasicColumn implements ISPColumn {
    public Name: string;
    constructor(name: string) {
        this.Name = name;
    }
}

export class LookupSPColumn implements ISPColumn {
    public LookupColumnName: string;
    public ProjectedColumns: Array<string>;
    public ExpandColumnSuffix: string = "";

    constructor(lookupColumnName: string, projectedColumns: Array<string>) {
        this.LookupColumnName = lookupColumnName;
        this.ProjectedColumns = projectedColumns;
    }
}

export class PeopleSPColumn extends LookupSPColumn {
    constructor(columnName: string, projectedColumns: Array<string>) {
        super(columnName, projectedColumns);
        this.ExpandColumnSuffix = "/Id";
    }
}

/* Entidades para Limit */

export interface ISPListItemLimit {
    LimitNumber: number;
}

export class SPListItemLimit implements ISPListItemLimit {
    public LimitNumber: number;
    /**
    * Crea un filtro por columna básico
    * @param LimitNumber Número de Items que se quiere recoger
    */
    constructor(limitNumber: number) {
        this.LimitNumber = limitNumber;
    }
}

/* Entidades para Order */
export enum SPOrderType {
    Desc, Asc
}

export interface ISPListItemOrder {
    ColumnName: string;
    OrderType: boolean;
}

export class SPListItemOrder implements ISPListItemOrder {
    public ColumnName: string;
    public OrderType: boolean;
    /**
    * Crea un filtro por columna básico
    * @param Nombre de la columna por la que filtrar
    * @param OrderType Dirección ascendente o descendente (asc/desc)
    */
    constructor(columnName: string, orderType: SPOrderType) {
        this.ColumnName = columnName;
        this.OrderType = (orderType == SPOrderType.Asc);
    }
}

/* Entitdades para Filter */

export enum SPFilterType {
    eq, lt, le, gt, ge, ne
}

export enum SPFilterOperator {
    and,
    or
}

export interface ISPListItemFilter {
    ColumnName: string;
    FilterType: SPFilterType;
    Value: string;
    FilterOperator: SPFilterOperator;
}

export class SPListItemFilter implements ISPListItemFilter {
    public ColumnName: string;
    public FilterType: SPFilterType;
    public Value: string;
    public FilterOperator: SPFilterOperator;
    /**
    * Crea un filtro por columna básico
    * @param Nombre de la columna por la que filtrar
    * @param filterType Tipo de filtro a aplicar (igual, mayor que, etc.)
    * @param filterValue Valor del filtro
    * @param filterOperator Valor del operador a aplicar en el filtro
    */
    constructor(columnName: string, filterType: SPFilterType, filterValue: string, filterOperator?: SPFilterOperator ) {
        this.ColumnName = columnName;
        this.FilterType = filterType;
        this.Value = filterValue;
        this.FilterOperator = filterOperator != null ? filterOperator : SPFilterOperator.and;
    }
}

export class SPListLookupItemFilter extends SPListItemFilter {
    public LookupColumnName: string;

    /**
    * Crea un filtro por columna de tipo lookup
    * @param columnName Nombre de la columna de tipo lookup
    * @param lookupListColumnName Nombre de la columna de la lista a la que apunta el lookup, por la que se filtrará
    * @param filterType Tipo de filtro a aplicar (igual, mayor que, etc.)
    * @param filterValue Valor del filtro
    * @param filterOperator Valor del operador a aplicar en el filtro
    */
    constructor(columnName: string, lookupListColumnName: string, filterType: SPFilterType, filterValue: string, filterOperator?: SPFilterOperator) {
        super(columnName, filterType, filterValue, filterOperator);
        this.LookupColumnName = lookupListColumnName;
    }
}