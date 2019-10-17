export class ArrayHelper {
    public static FirstOrDefault<T>(array: Array<T>, compareFunction?: (elementOfArray: T, indexInArray: string) => boolean) {
        if (!array) return null;
        let match: T = null;
        if (compareFunction) {
            for (let index in array) {
                let item = array[index];
                if (compareFunction(item, index)) {
                    match = item;
                    break;
                }
            }
        }
        return match;
    }

    public static Select<T, J>(array: Array<T>, selectFunction: (elementOfArray: T, index?: number) => J): Array<J> {
        if (!array || array.length == 0) return [];
        let responseArray: Array<J> = [];
        array.forEach((item, index) => {
            responseArray.push(selectFunction(item, index));
        });
        return responseArray;
    }

    public static SelectMany<T, J>(array: Array<T>, selectFunction: (elementOfArray: T, index?: number) => Array<J>): Array<J> {
        if (!array || array.length == 0) return [];
        let responseArray: Array<J> = [];
        array.forEach((item, index) => {
            let selectItems = selectFunction(item, index);
            selectItems.forEach(selectItem => {
                responseArray.push(selectItem);
            });
        });
        return responseArray;
    }

    public static Where<T>(array: Array<T>, compareFunction: (elementOfArray: T) => boolean): Array<T> {
        if (!array || array.length == 0) return [];
        let responseArray: Array<T> = [];
        array.forEach(item => {
            if (compareFunction(item))
                responseArray.push(item);
        });
        return responseArray;
    }

    public static Any<T>(array: Array<T>, compareFunction?: (elementOfArray: T) => boolean): boolean {
        if (!array || array.length == 0) return false;
        if (compareFunction == null) return true;
        let found = false;
        for (let item of array) {
            if (compareFunction(item)) {
                found = true;
                break;
            }
        }
        return found;
    }

    public static ConvertToDictionary<T>(array: Array<T>, keyFunction: (elementOfArray: T) => string): { [key: string]: T } {
        if (!array) return null;
        var resultDict: { [key: string]: T } = {};
        array.forEach(item => {
            let key = keyFunction(item);
            resultDict[key] = item;
        });
        return resultDict;
    }

    public static RemoveFirstMatch<T>(array: Array<T>, matchFunction: (elementOfArray: T) => boolean): boolean {
        if (!array) return false;
        for (let i = 0; i < array.length; i++) {
            let matches = matchFunction(array[i]);
            if (matches) {
                array.splice(i, 1);
                return true;
            }
        }
        return false;
    }

    public static GroupBy<T, J>(array: Array<T>, keyFunction: (elementOfArray: T) => string, selectFunction: (item: T) => J): { [key: string]: Array<J> } {
        if (!array) return null;
        let resultDict: { [key: string]: Array<J> } = {};
        array.forEach(item => {
            var key = keyFunction(item);
            if (!resultDict[key]) resultDict[key] = [];
            resultDict[key].push(selectFunction(item));
        });
        return resultDict;
    }

    public static SortBy<T>(array: Array<T>, keyFunction: (elementOfArray: T) => string): Array<T> {
        if (!array) return null;
        return array.sort((a: T, b: T) => {
            var aKey = keyFunction(a);
            var bKey = keyFunction(b);
            if (aKey < bKey) return -1;
            else if (aKey > bKey) return 1;
            return 0;
        });
    }

}