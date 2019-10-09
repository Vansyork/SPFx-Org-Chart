import { ValueEntity } from "../interfaces/IGraphUserdata";
import { IPerson } from "../interfaces/IPerson";
import { IPersonListItem } from "../interfaces/IPersonListItem";
export class Person implements IPerson {
    public id: string | number;
    public name: string;
    public department: string;
    public description?: string;
    public imageUrl?: string;
    public children?: Person[] = [];
    public originalListItem: IPersonListItem;
    constructor(listItem: IPersonListItem, allUsersData: IPersonListItem[] | ValueEntity[]) {
        this.id = listItem.Id;
        this.name = listItem.Title;
        this.department = listItem.ORG_Department;
        this.description = listItem.ORG_Description || null;
        this.imageUrl = listItem.ORG_Picture ? listItem.ORG_Picture.Url : null;
        this.originalListItem = listItem;

        listItem.ORG_MyReportees.forEach(reportee => {
            //u.Id != this.id filter wrongly configured user as reportee
            if (this.isValueEntity(allUsersData)) {
                let filterdUserData: ValueEntity[];
                filterdUserData = (allUsersData as ValueEntity[]).filter((u) => { return (u.id === reportee.Id && u.id != this.id); });
                if (filterdUserData.length > 0) {
                    filterdUserData.forEach((user: ValueEntity) => {
                        let listItem: IPersonListItem = {
                            Id: user.id,
                            Title: user.displayName,
                            ORG_Department: user.jobTitle,
                            ORG_Description: user.jobTitle,
                            ORG_Picture: { Url: null },
                            // ORG_MyReportees: (allUsersData as ValueEntity[]).map((val: ValueEntity) => { return { Id: val.id }; })
                            ORG_MyReportees: []
                        };
                        this.children.push(new Person(listItem, allUsersData));
                    });
                }
            } else {
                let filterdUserData: IPersonListItem[];
                filterdUserData = (allUsersData as IPersonListItem[]).filter((u) => { return (u.Id === reportee.Id && u.Id != this.id); });
                if (filterdUserData.length > 0) {
                    filterdUserData.forEach(element => {
                        this.children.push(new Person(element, allUsersData));
                    });
                }
            }
        });

    }

    private isValueEntity(arg: Array<any>): arg is ValueEntity[] {
        if (arg.length > 1) {
            return arg[0].displayName !== undefined;
        } else {
            return false;
        }
    }
}