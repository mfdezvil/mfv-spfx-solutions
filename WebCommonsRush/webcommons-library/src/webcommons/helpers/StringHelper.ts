export class StringHelper {
    
    public static TrimChar (originalString: string, char: string) {
        var regEx = new RegExp("(^" + char.charAt(0) + ")|(" + char.charAt(0) + "$)");
        return originalString.replace(regEx, "");
    }

    public static Contains (originalString: string, stringToFind: string) {
        return originalString.indexOf(stringToFind) != -1;
    }

    public static SplitString(originalString: string, substringToSplit:string, returnLastPart?:boolean) {
        if(StringHelper.Contains(originalString,substringToSplit)) {
            let splitAux= originalString.split(substringToSplit);
            if(!returnLastPart)
                return splitAux[0];
            return splitAux[1];
        }
        return originalString;
    }

}