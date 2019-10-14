import { IGraphUserdata, ValueEntity } from "../interfaces/IGraphUserdata";
import { IPerson } from "../interfaces/IPerson";
import { IPersonListItem, IReportee } from "../interfaces/IPersonListItem";
import DataService from "../services/dataservice";
export class Person implements IPerson {
    public id: string | number;
    public name: string;
    public children: Person[] = [];
    public department?: string;
    public description?: string;
    public imageUrl?: string;
    public email?: string;
    constructor(listItem: IPersonListItem, allUsersData?: IPersonListItem[], dataService?: DataService, setStateFunc?: Function) {
        this.id = listItem.Id;
        this.name = listItem.Title;
        this.department = listItem.ORG_Department;
        this.description = listItem.ORG_Description || null;
        this.imageUrl = listItem.ORG_Picture ? listItem.ORG_Picture.Url : null;
        this.email = listItem.email ? listItem.email : null;

        if (allUsersData) {
            listItem.ORG_MyReportees.forEach((reportee: IReportee) => {
                //u.Id != this.id filter wrongly configured user as reportee
                let filterdUserData: IPersonListItem[];
                filterdUserData = allUsersData.filter((u) => { return (u.Id === reportee.Id && u.Id != this.id); });
                if (filterdUserData.length > 0) {
                    filterdUserData.forEach(element => {
                        this.children.push(new Person(element, allUsersData));
                    });
                }
            });
        } else if (this.email) {
            dataService.getDirectReportsForUserFromGraphAPI(this.email).then(
                (result: IGraphUserdata) => {
                    result.value.forEach((element: ValueEntity) => {
                        this.children.push(
                            new Person(
                                {
                                    Id: element.id,
                                    Title: element.displayName,
                                    email: element.mail
                                }, null, dataService, setStateFunc));
                    });
                    if (setStateFunc) {
                        setStateFunc();
                    }
                }
            );
        }
    }

}