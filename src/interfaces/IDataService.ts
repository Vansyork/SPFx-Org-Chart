import { IGraphUserdata } from "./IGraphUserdata";
import { IList } from "./IList";
import { IPerson } from './IPerson';
import { IPersonListItem } from "./IPersonListItem";


export interface IDataService {
    getDirectReportsForUser(list: string, user: string): Promise<IPerson>;
    getDirectReportsForUserFromGraphAPI(email: string): Promise<IGraphUserdata>;
    getUserPhotoFromGraphApi(userEmail: string);
    getUserInfoFromGraphApi(userEmail: string);
    getOrgList(): Promise<IList[]>;
    getUsersFromList(listid: string): Promise<IPersonListItem[]>;
    checkIfListAlreadyExists(listName: string): Promise<boolean>;
    createList(listName: string): Promise<IList>;
}