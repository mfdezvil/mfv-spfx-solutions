import { IUserItem } from "../models/IUserItem";
import {Utils} from "../utils/Utils";

export class CurrentUserDataHelper {
    public static GetUserFromResponse(response): IUserItem {
        var user: IUserItem = {
            displayName: response.displayName,
            mail: (!Utils.IsNullOrEmpty(response.mail)) ? response.mail : null,
            userPrincipalName: (!Utils.IsNullOrEmpty(response.userPrincipalName)) ? response.userPrincipalName : null
          };
        return user;
    }



}