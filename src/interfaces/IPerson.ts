
export interface IPerson {
    id: number | string;
    name: string;
    department?: string;
    description?: string;
    imageUrl?: string;
    children?: IPerson[];
    email?: string;
}