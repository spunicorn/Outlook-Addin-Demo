export interface IHttpClient{
    fetch<T>(url,body) : Promise<T>;
}