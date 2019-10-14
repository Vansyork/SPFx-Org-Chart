import { IDataService } from '../interfaces/IDataService';
import { IGraphUserdata } from '../interfaces/IGraphUserdata';
import { IList } from "../interfaces/IList";
import { IPerson } from '../interfaces/IPerson';
import { IPersonListItem } from "../interfaces/IPersonListItem";

export default class MockDataService implements IDataService {
  public getUserPhotoFromGraphApi(userEmail: string) {
    return Promise.reject("Image not found");
  }
  public getUserInfoFromGraphApi(userEmail: string) {
    return Promise.resolve({ jobTitle: "Developer", department: "IT" });
  }
  public getDirectReportsForUserFromGraphAPI(email: string): Promise<IGraphUserdata> {

    let results = [
      { "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#directoryObjects", "value": [{ "@odata.type": "#microsoft.graph.user", "id": "685da7e6-e430-4961-8d95-e673ce1eecdc", "businessPhones": ["123.456.789"], "displayName": "Peppie en Kokki", "givenName": "Peppie", "jobTitle": "Developer", "mail": "Peppie.Kokki@contoso.onmicrosoft.com", "mobilePhone": null, "officeLocation": "Kalmthout", "preferredLanguage": null, "surname": "Kokki", "userPrincipalName": "Peppie.Kokki@contoso.onmicrosoft.com" }, { "@odata.type": "#microsoft.graph.user", "id": "808534ae-9ee9-4480-a7cd-266df72d3f6e", "businessPhones": [], "displayName": "Gianluigi  Buffon", "givenName": "Gianluigi", "jobTitle": "Goalkeeper", "mail": "Gianluigi.Buffon@contoso.onmicrosoft.com", "mobilePhone": null, "officeLocation": "Torino", "preferredLanguage": null, "surname": "Buffon", "userPrincipalName": "Gianluigi.Buffon@contoso.onmicrosoft.com" }] }
      , { "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#directoryObjects", "value": [] }
      , { "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#directoryObjects", "value": [{ "@odata.type": "#microsoft.graph.user", "id": "45b03962-66f5-484f-9252-c4745b35dd83", "businessPhones": [], "displayName": "Kevin Kkorn", "givenName": "Kevin", "jobTitle": "Sales", "mail": "Kevin.Kkorn@contoso.onmicrosoft.com", "mobilePhone": null, "officeLocation": "Diegem", "preferredLanguage": null, "surname": "Kkorn", "userPrincipalName": "Kevin.Kkorn@contoso.onmicrosoft.com" }] }
      , { "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#directoryObjects", "value": [{ "@odata.type": "#microsoft.graph.user", "id": "01cbc81f-f13b-41ff-94ac-e17723836a98", "businessPhones": [], "displayName": "Big Jokke", "givenName": "Big", "jobTitle": "Support", "mail": null, "mobilePhone": null, "officeLocation": null, "preferredLanguage": null, "surname": "Jokke", "userPrincipalName": "Big.Jokke@contoso.onmicrosoft.com" }] }];

    return Promise.resolve(results[Math.floor(Math.random() * 4)]);
  }
  public checkIfListAlreadyExists(listName: string): Promise<boolean> {
    return Promise.resolve(false);
  }
  public createList(listName: string): Promise<IList> {
    return Promise.resolve({ Id: 0, Title: listName, ParentWebUrl: "/sites/contoso" });
  }
  public getUsersFromList(listid: string): Promise<IPersonListItem[]> {
    let result: IPersonListItem[] = [{
      Id: "1",
      Title: "Vans York",
      ORG_Department: "SharePoint Team",
      ORG_Description: "The One and Only",
      ORG_Picture: { Url: "https://contoso.sharepoint.com/:i:/s/000334/EU3ZLIxyrypOuZruJLKfE9UB_C94kBzaMdsYXCun1WzQtw?e=3chAdd" },
      ORG_MyReportees: [{ Id: '2' }]
    },
    {
      Id: "2",
      Title: "That Man",
      ORG_Department: "SharePoint Team",
      ORG_Description: "SharePoint technical lead",
      ORG_Picture: null,
      ORG_MyReportees: []
    }];
    return Promise.resolve(result);
  }
  public getOrgList(): Promise<IList[]> {
    let result: IList[] = [
      { Id: "1", Title: "Org List 1" },
      { Id: "2", Title: "Org List 2" },
      { Id: "3", Title: "Org List 3" }
    ];
    return Promise.resolve(result);
  }
  public getDirectReportsForUser(listid: string, userid: string): Promise<IPerson> {

    var initechOrg: IPerson = {
      "children": [{
        "children": [{
          "children": [{
            "children": [],
            "id": 3,
            "name": "That Man",
            "department": "SharePoint Team",
            "description": "Technical lead",
            "imageUrl": null
          },
          {
            "children": [{
              "children": [],
              "id": 2,
              "name": "Guy Big",
              "department": "SharePoint Team",
              "description": "Support envangelist",
              "imageUrl": null
            }],
            "id": 1,
            "name": "Man the Genious",
            "department": "SharePoint Team",
            "description": "The One and Only",
            "imageUrl": "https://contoso.sharepoint.com/sites/000334/SiteAssets/IMG_20170505_160501.jpg"
          }],
          "id": 4,
          "name": "Would be Guy",
          "department": "SharePoint Team",
          "description": "Team leader",
          "imageUrl": null
        },
        {
          "children": [{
            "children": [],
            "id": 7,
            "name": "Social Guy",
            "department": "IAP",
            "description": null,
            "imageUrl": null
          },
          {
            "children": [],
            "id": 6,
            "name": "The Girl",
            "department": "IAP",
            "description": "Head IAP",
            "imageUrl": null
          }],
          "id": 8,
          "name": "Good Guy L.",
          "department": "IAP",
          "description": null,
          "imageUrl": null
        }],
        "id": 5,
        "name": "Nice Man",
        "department": "Head Applications",
        "description": "The big Boss",
        "imageUrl": null
      },
      {
        "children": [{
          "children": [{
            "children": [{
              "children": [],
              "id": 3,
              "name": "That Man",
              "department": "SharePoint Team",
              "description": "Technical lead",
              "imageUrl": null
            },
            {
              "children": [{
                "children": [],
                "id": 2,
                "name": "Guy Big",
                "department": "SharePoint Team",
                "description": "Support envangelist",
                "imageUrl": null
              }],
              "id": 1,
              "name": "Man the Genious",
              "department": "SharePoint Team",
              "description": "The One and Only",
              "imageUrl": "https://contoso.sharepoint.com/sites/000334/SiteAssets/IMG_20170505_160501.jpg"
            }],
            "id": 4,
            "name": "Would be Guy",
            "department": "SharePoint Team",
            "description": "Team leader",
            "imageUrl": null
          },
          {
            "children": [{
              "children": [],
              "id": 7,
              "name": "Social Guy",
              "department": "IAP",
              "description": null,
              "imageUrl": null
            },
            {
              "children": [],
              "id": 6,
              "name": "The Girl",
              "department": "IAP",
              "description": "Head IAP",
              "imageUrl": null
            }],
            "id": 8,
            "name": "Good Guy L.",
            "department": "IAP",
            "description": null,
            "imageUrl": null
          }],
          "id": 5,
          "name": "Nice Man",
          "department": "Head Applications",
          "description": "The big Boss",
          "imageUrl": null
        }],
        "id": 10,
        "name": "Rope Man",
        "department": "Business Partners",
        "description": null,
        "imageUrl": null
      }],
      "id": 9,
      "name": "Vans York ",
      "department": "IT Head",
      "description": null,
      "imageUrl": null
    };

    var initechOrg2: IPerson = {
      "children": [{
        "id": 10,
        "name": "Rope Man",
        "department": "Business Partners",
        "description": null,
        "imageUrl": null
      }],
      "id": 9,
      "name": "Vans York ",
      "department": "IT Head",
      "description": null,
      "imageUrl": null
    };

    return Promise.resolve(listid === "2" ? initechOrg2 : initechOrg);
  }
}