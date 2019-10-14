
export interface IPersonListItem {
    Id: number | string;
    Title: string;
    ORG_Department?: string;
    ORG_Description?: string;
    ORG_Picture?: IPicture;
    ORG_MyReportees?: IReportee[];
    email?: string;
}

export interface IPicture {
    Url: string;
}

export interface IReportee {
    Id: string | number;
}