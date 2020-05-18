import { IHttpClient } from "./IHttpClient";
import { IAuthorization } from "./auth/IAuthorization";

export class FetchHttpClient implements IHttpClient {
    constructor(protected authorization?: IAuthorization){

    }
    public async fetch<T>(url: any, options: any): Promise<T> {
        if(this.authorization){
            let authHeader = await this.authorization.authorize(url);
            options.headers = options.headers || new Headers();
            options.headers.append([authHeader.key], authHeader.value);
        }
        let fetchResponse = await fetch(url, options);
        return fetchResponse.json();
    }
}