import { IGraphDataReader } from "../../managers/IGraphDataReader";
import { ISPGraphUser } from "../../entities/ISPGraphUser";
import { ISPGraphMails } from "../../entities/ISPGraphMail";
import { ISPGraphTeams } from "../../entities/ISPGraphTeam";

export class GraphMock implements IGraphDataReader {

    private static userInfo: ISPGraphUser = {
        displayName: "Maria Fernandez Villanueva",
        mail: "mariaf@mfvcyc2.onmicrosoft.com"
    };

    private static userMail: ISPGraphMails = {
        value: [
            {
                id: "AAMkADBhMTE0MzJjLWQyM2ItNDgzNS1hZGI2LWNlY2Q1ZTA1MDFkMgBGAAAAAADuI5HmMdHBRpfJAW5lvvphBwD67V4LXN-VQZQFjtSyrf_JAAAAAAEMAAD67V4LXN-VQZQFjtSyrf_JAAAdrcR4AAA=",
                receivedDateTime: "2019-04-25T10:25:33Z",
                subject: "Información de cuenta para usuarios nuevos o modificados",
                from: {
                    emailAddress: {
                        address: "ms-noreply@microsoft.com",
                        name: "Microsoft on behalf of your organization"
                    }
                },
                isRead: false
            }
        ]
    };

    private static userEvents: any = {
        value: [
            {
                "id": "AAMkADBhMTE0MzJjLWQyM2ItNDgzNS1hZGI2LWNlY2Q1ZTA1MDFkMgBGAAAAAADuI5HmMdHBRpfJAW5lvvphBwD67V4LXN-VQZQFjtSyrf_JAAAAAAENAAD67V4LXN-VQZQFjtSyrf_JAAAdrfhUAAA=",
                "subject": "Reunión muy importante",
                "bodyPreview": "",
                "body": {
                    "contentType": "html",
                    "content": "<html>\r\n<head>\r\n<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\">\r\n<meta content=\"text/html; charset=iso-8859-2\">\r\n<style type=\"text/css\" style=\"display:none\">\r\n<!--\r\np\r\n\t{margin-top:0;\r\n\tmargin-bottom:0}\r\n-->\r\n</style>\r\n</head>\r\n<body dir=\"ltr\">\r\n<div id=\"divtagdefaultwrapper\" dir=\"ltr\" style=\"font-size:12pt; color:#000000; font-family:Calibri,Helvetica,sans-serif\">\r\n<p style=\"margin-top:0; margin-bottom:0\"><br>\r\n</p>\r\n</div>\r\n</body>\r\n</html>\r\n"
                },
                "start": {
                    "dateTime": "2019-05-13T08:00:00.0000000",
                    "timeZone": "UTC"
                },
                "end": {
                    "dateTime": "2019-05-13T12:00:00.0000000",
                    "timeZone": "UTC"
                },
                "location": {
                    "displayName": "Calle de Velázquez, 117",
                    "locationUri": "https://www.bingapis.com/api/v6/addresses/QWRkcmVzcy8tMjk0NDQ1NzkzOSU3YzExNz9hbHRRdWVyeT1hbCU1ZUNhbGxlK2RlK1ZlbCVjMyVhMXpxdWV6JTJjKzExNyU3Y2xjJTVlTWFkcmlkJTdjYTElNWVDb211bmlkYWQrZGUrTWFkcmlkJTdjY3IlNWVFc3BhJWMzJWIxYSU3Y2lzbyU1ZUVT?setLang=es",
                    "locationType": "streetAddress",
                    "uniqueId": "https://www.bingapis.com/api/v6/addresses/QWRkcmVzcy8tMjk0NDQ1NzkzOSU3YzExNz9hbHRRdWVyeT1hbCU1ZUNhbGxlK2RlK1ZlbCVjMyVhMXpxdWV6JTJjKzExNyU3Y2xjJTVlTWFkcmlkJTdjYTElNWVDb211bmlkYWQrZGUrTWFkcmlkJTdjY3IlNWVFc3BhJWMzJWIxYSU3Y2lzbyU1ZUVT?setLang=es",
                    "uniqueIdType": "bing",
                    "address": {
                        "street": "Calle de Velázquez, 117",
                        "city": "Madrid",
                        "state": "Comunidad de Madrid",
                        "countryOrRegion": "España",
                        "postalCode": "28006"
                    },
                    "coordinates": {
                        "latitude": 40.436981201171875,
                        "longitude": -3.683340072631836
                    }
                },
                "attendees": [],
                "organizer": {
                    "emailAddress": {
                        "name": "Maria Fernandez Villanueva",
                        "address": "mariaf@mfvcyc2.onmicrosoft.com"
                    }
                }
            }
        ]
    };

    private static userTeams: any = {
        value: [
            {id: "a693d445-6a76-4c8e-a78d-48d0919addd9", displayName: "Equipo público 1", description: "Este es un equipo público"},
            {id: "b9238e80-875e-41a4-9af2-9e987559d640", displayName: "Equipo público 2", description: "Este es el equipo público 2"},
            {id: "aa051256-d889-48b7-8be0-9fc3450433b9", displayName: "Equipo privado 1", description: "Este es un equipo privado"}
        ]
    };

    constructor() {}

    public GetCurrentUserData(): Promise<ISPGraphUser> {
        return new Promise<ISPGraphUser>((resolve) => {
            resolve(GraphMock.userInfo);
        });
    }

    public GetCurrentUserMails(): Promise<ISPGraphMails> {
        return new Promise<ISPGraphMails>((resolve) => {
            resolve(GraphMock.userMail);
        });
    }

    public GetCurrentUserEvents(): Promise<any> {
        return new Promise<any>((resolve) => {
            resolve(GraphMock.userEvents);
        });
    }

    public GetCurrentUserTeams(): Promise<ISPGraphTeams> {
        return new Promise<ISPGraphTeams>((resolve) => {
            resolve(GraphMock.userTeams);
        });
    }

}