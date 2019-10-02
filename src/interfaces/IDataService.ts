import { IPerson } from './IPerson';
import { IList } from "./IList";
import { IPersonListItem } from "./IPersonListItem";

import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IDataService {
    getDirectReportsForUser(list:string, user: string): Promise<IPerson>;
    getOrgList(): Promise<IList[]>;
    getUsersFromList(listid: string): Promise<IPersonListItem[]>;
    checkIfListAlreadyExists(listName: string): Promise<boolean>;
    createList(listName: string): Promise<IList>;
}