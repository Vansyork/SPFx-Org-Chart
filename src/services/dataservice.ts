import { MSGraphClient, SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import pnp, { ContentTypeAddResult, List, ListAddResult, ProcessHttpClientResponseException } from "@pnp/pnpjs";
import ErrorHandler from '../helpers/ErrorHandler';
import { IDataService } from '../interfaces/IDataService';
import { IGraphUserdata } from "../interfaces/IGraphUserdata";
import { IList } from "../interfaces/IList";
import { IPerson } from '../interfaces/IPerson';
import { IPersonListItem } from "../interfaces/IPersonListItem";
import { SPContentType } from '../interfaces/SPContentType';
import { SPListData } from "../interfaces/SPListData";
import { Person } from '../models/person';

export default class DataService implements IDataService {
  constructor(protected context: WebPartContext) { }

  //#region public methods
  public checkIfListAlreadyExists(listName: string): Promise<boolean> {
    return pnp.sp.web.lists.getByTitle(listName).get().then((listResult: List) => {
      if (listResult) {
        return Promise.resolve(true);
      }
    })
      .catch((e: ProcessHttpClientResponseException) => {
        if (e.status === 404) {
          return Promise.resolve(false);
        }
        else {
          return ErrorHandler.handleError(e);
        }
      });
  }
  public createList(listName: string): Promise<IList> {
    return pnp.sp.web.lists.add(listName, "List to configure the org chart webpart", 100, true).then((orgListAddResult: ListAddResult) => {
      return this.configureOrgList((orgListAddResult)).then(() => {
        return pnp.sp.web.lists.getById(orgListAddResult.data.Id).views.get().then((views: any[]) => {
          let defaultView: any = views.filter((v) => { return v.DefaultView === true; }).shift();
          return Promise.resolve(<IList>{ Id: orgListAddResult.data.Id, Title: orgListAddResult.data.Title, ParentWebUrl: orgListAddResult.data.ParentWebUrl, NavUrl: defaultView.ServerRelativeUrl });
        });
      });
    }).catch(ErrorHandler.handleError);
  }

  public getUsersFromList(listid: string): Promise<IPersonListItem[]> {
    return this.context.spHttpClient.get(
      `${this.context.pageContext.web.absoluteUrl}/_api/web/lists('${listid}')/items?$select=Id,Title,ORG_Department,ORG_Description,ORG_Picture,ORG_MyReportees,ORG_MyReportees/Id&$expand=ORG_MyReportees`,
      SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => response.json())
      .then((jsonData: { value: IPersonListItem[] }) => {
        return jsonData.value;
      }).catch(ErrorHandler.handleError);
  }

  public getOrgList(): Promise<IList[]> {
    return this.context.spHttpClient.get(
      `${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$filter=Hidden eq false and BaseTemplate eq 100&$select=Id,Title`,
      SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => response.json())
      .then((jsonData: { value: IList[] }) => {
        return this.filterOrgChartContentTypesFromLists(jsonData.value);
      }).catch(ErrorHandler.handleError);
  }

  public getDirectReportsForUser(listid: string, userid: string): Promise<IPerson> {
    return this.getUsersFromList(listid).then((users: IPersonListItem[]) => {
      let filteredArray: Person[] = users.filter((u: IPersonListItem) => { return u.Id === userid; }).map(
        (filteredUser: IPersonListItem) => {
          let startUser = new Person(filteredUser, users);
          return startUser;
        });

      if (filteredArray.length === 1) {
        return Promise.resolve(filteredArray[0]);
      }
      else {
        return ErrorHandler.handleError("error getting direct reports for user");
      }
    }).catch(ErrorHandler.handleError);
  }

  public getDirectReportsForUserFromGraphAPI(userEmail: string): Promise<IGraphUserdata> {
    return this.context.msGraphClientFactory.getClient()
      .then((client: MSGraphClient) => {
        return client
          .api(`users/${userEmail}/directReports`)
          .version("v1.0")
          .get()
          .catch(ErrorHandler.handleError);
      });
  }

  public getUserPhotoFromGraphApi(userEmail: string) {
    return this.context.msGraphClientFactory.getClient()
      .then((client: MSGraphClient) => {
        return client
          .api(`users/${userEmail}/photo/$value`)
          .version('v1.0')
          .responseType('blob')
          .get();
      });
  }

  public getUserInfoFromGraphApi(userEmail: string) {
    return this.context.msGraphClientFactory.getClient()
      .then((client: MSGraphClient) => {
        return client
          .api(`users/${userEmail}`)
          .version('beta')
          .get();
      });
  }
  //#endregion

  //#region private methods

  private configureOrgList(spListAddResult: ListAddResult): Promise<void> {
    return spListAddResult.list.contentTypes.addAvailableContentType("0x0100F4C266967DF54F5FAB9CDAA2A09D51C9").then((addedCTResult: ContentTypeAddResult) => {
      return this.getItemCT(spListAddResult).then((contentype: SPContentType) => {
        return this.updateLookupField(spListAddResult.list, spListAddResult.data).then(() => {
          return this.updateView(spListAddResult.list).then(() => {
            return spListAddResult.list.contentTypes.getById(contentype.StringId).delete().then(() => {
              return Promise.resolve();
            });
          });
        });
      });
    }).catch(ErrorHandler.handleError);
  }

  private updateView(spList: List): Promise<void> {
    let batch = pnp.sp.createBatch();

    spList.views.getByTitle("All Items").fields.inBatch(batch).add("ORG_Department"),
      spList.views.getByTitle("All Items").fields.inBatch(batch).add("ORG_Description"),
      spList.views.getByTitle("All Items").fields.inBatch(batch).add("ORG_Picture"),
      spList.views.getByTitle("All Items").fields.inBatch(batch).add("ORG_MyReportees"),
      spList.views.getByTitle("All Items").fields.inBatch(batch).add("ORG_MyReportees_ID");

    return batch.execute().then(() => {
      return Promise.resolve();
    }).catch(ErrorHandler.handleError);
  }

  private updateLookupField(spList: List, listData: SPListData): Promise<void> {
    return spList.fields.getByInternalNameOrTitle("ORG_MyReportees").update({
      "SchemaXml":
        `<Field Type="LookupMulti"
              DisplayName="My Reportees"
              Required="FALSE"
              List="{${listData.Id}}"
              EnforceUniqueValues="FALSE"
              ShowField="Title"
              Mult="TRUE"
              Sortable="FALSE"
              UnlimitedLengthInDocumentLibrary="FALSE"
              RelationshipDeleteBehavior="None"
              ID="{F84FC9D9-6307-44BA-84C5-C029C0D19BE8}"
              StaticName="ORG_MyReportees"
              Name="ORG_MyReportees"
              Group="ORG Columns" />`
    }).then(() => {
      return Promise.resolve();
    }).catch(ErrorHandler.handleError);
  }

  private getItemCT(spListAddResult: ListAddResult): Promise<SPContentType> {
    return spListAddResult.list.contentTypes.get().then((contentTypes: SPContentType[]) => {
      let filteredCTs = contentTypes.filter((ct) => ct.Name === "Item");
      if (filteredCTs.length === 1) {
        return Promise.resolve(filteredCTs[0]);
      }
      else {
        return Promise.reject("Could not return Item contenttype");
      }
    }).catch(ErrorHandler.handleError);
  }

  private filterOrgChartContentTypesFromLists(Lists: IList[]): Promise<IList[]> {
    let batch = pnp.sp.createBatch();
    let filteredLists: IList[] = [];
    Lists.forEach(lst => {
      pnp.sp.web.lists.getById(lst.Id.toString()).contentTypes.inBatch(batch).get().then((cts: [{ Id: { StringValue: String } }]) => {
        cts.forEach(contentType => {
          if (contentType.Id.StringValue.indexOf("0x0100F4C266967DF54F5FAB9CDAA2A09D51C9") !== -1) {
            filteredLists.push(lst);
          }
        });
      }, ErrorHandler.handleError);
    });
    return batch.execute().then(() => {
      return filteredLists;
    });
  }
  //#endregion
}