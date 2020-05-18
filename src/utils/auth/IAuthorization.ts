export interface IAuthorization{
    authorize(resource):Promise<{key:string, value: string}>;
}