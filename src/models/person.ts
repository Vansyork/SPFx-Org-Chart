import { IPerson } from "../interfaces/IPerson";
import { IPersonListItem } from "../interfaces/IPersonListItem";
export class Person implements IPerson {
    public id: string | number;
    public name: string;
    public department: string;
    public description?: string;
    public imageUrl?: string;
    public children?: Person[] = [];
    constructor(listItem: IPersonListItem, allUsersData: IPersonListItem[]) {
        this.id = listItem.Id;
        this.name = listItem.Title;
        this.department = listItem.ORG_Department;
        this.description = listItem.ORG_Description || null;
        this.imageUrl = listItem.ORG_Picture ? listItem.ORG_Picture.Url : null;

        listItem.ORG_MyReportees.forEach(reportee => {
            let userdata: IPersonListItem;
            let filterdUserData: IPersonListItem[];
            //u.Id != this.id filter wrongly configured user as reportee
            filterdUserData = allUsersData.filter((u) => { return (u.Id === reportee.Id && u.Id != this.id); });
            if (filterdUserData.length > 0) {
                filterdUserData.forEach(element => {
                    this.children.push(new Person(element, allUsersData));
                });
            }
        });

    }
}



