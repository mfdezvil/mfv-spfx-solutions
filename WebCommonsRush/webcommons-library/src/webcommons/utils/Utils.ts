export class Utils { 

    public static IsNullOrEmpty(object: any) {
        if (object == undefined || object == null) return true;
        if (object.constructor === Object) return Object.keys(object).length === 0;
        return object == "";
    }
}