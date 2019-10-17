import {
    ClientSideText, 
    ClientSideWebpart,
    sp,
    ClientSidePage,
    ItemUpdateResult
} from "@pnp/sp";
import { INewsPage } from "../models/INewsPage";

export class PageHelper {

    public static async createCustomPage(newsPageData: INewsPage): Promise<string> {
        const page = await sp.web.addClientSidePage(newsPageData.newPageName + ".aspx");
             
        const partDefs = await sp.web.getClientSideWebParts();
        const section = page.addSection();
        console.log("section added");

        const column1 = section.addColumn(4);

        // find the definition we want, here by id
        const partDef = partDefs.filter(c => c.Id === "7f718435-ee4d-431c-bdbf-9c4ff326f46e");

        // optionally ensure you found the def
        if (partDef.length < 1) {
            // we didn't find it so we throw an error
            console.log('ops');
            throw new Error("Could not find the web part");
        }
        // create a ClientWebPart instance from the definition
        const part = ClientSideWebpart.fromComponentDef(partDef[0]);
        /*
        part.setProperties<MyClientSideWebpartPropertyTypes.People>({
            layout: "2",
            persons: [
                {
                    "id": "i:0#.f|membership|jsmith@federicoporceddu.onmicrosoft.com",
                    "upn": "jsmith@federicoporceddu.onmicroosft.com",
                    "role": "",
                    "department": "",
                    "phone": "",
                    "sip": ""
                }
            ]
        });*/
        // add a text control to the second new column
        column1.addControl(part);

        const column2 = section.addColumn(8);
        //// add a text control to the second new column
        column2.addControl(new ClientSideText("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum."));

        page.disableComments();
        await page.save();
        console.log("case saved");

        var id = await sp.web.lists.getByTitle("Site Pages").items.filter("Title eq \'" + newsPageData.newPageName + "\'").select("Id").getAll(1);
        if(id.length > 0) {
            console.log("Calling update Page properties called for - " + id[0].Id);
            await sp.web.lists.getByTitle("Site Pages").items.getById(id[0].Id).update({
                CyCNewsType: newsPageData.newsPageType != "" ? newsPageData.newsPageType : null,
                Title : newsPageData.newPageName,
                CyCSiteDateExpiration: newsPageData.newsExpirationDate
            }).then((iar: ItemUpdateResult) => {
                console.log(iar);
            });
        }

        return await "done";
    }
}