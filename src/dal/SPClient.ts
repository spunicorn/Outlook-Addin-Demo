import { IHttpClient } from "../utils/IHttpClient";
import { MsalAuthorization } from "../utils/auth/MsalAuthorization";
import { FetchHttpClient } from "../utils/FetchHttpClient";
import { ICustomerRecord } from "../model/dto/ICustomerRecord";
export interface ISharePointNoMetadataResponse{
    value:ICustomerRecord[]
}


export class SPClient {
    protected HttpClient: IHttpClient;
    public constructor() {
        let authorization = new MsalAuthorization("e728e4c0-a694-4432-8624-64813827cf39",
            "https://localhost:3000/taskpane.html",
            "e370efe0-35be-4e74-979e-b57598b68d4f", {
            scopes: ["https://spunicorntest.sharepoint.com/.default"]
        });
        this.HttpClient = new FetchHttpClient(authorization);
    }
    public async getSites(): Promise<any> {
        return this.HttpClient.fetch("https://spunicorntest.sharepoint.com/_api/search/query?queryText='ContentClass:STS_Site'&selectProperties='Title,Id,Path'", {});
    }
    public async getProductsRelatedCustomer(email:string):Promise<ICustomerRecord[]>{
        let queryUrl = `https://spunicorntest.sharepoint.com/sites/testteamsite/_api/lists/getByTitle('CustomerInfo')/items?$filter=Email eq '${email}'`;
        let something = await this.HttpClient.fetch<ISharePointNoMetadataResponse>(queryUrl,{headers:{
                                'Content-Type': 'application/json',
                                'Accept': 'application/json;odata=nometadata'
                            }});
        return something.value;
        // return new Promise<ICustomerRecord[]>((resolve,reject)=>{
        //     this.HttpClient.fetch(queryUrl, {
        //         headers: {
        //         'Content-Type': 'application/json',
        //         'Accept': 'application/json;odata=nometadata',
        //     }}).then((response:Response ) =>{
        //         console.log(response)
        //         response.json().then((responseJson) =>{
        //             console.log(responseJson)
        //             if(responseJson.value){
        //                 resolve(responseJson.value as ICustomerRecord[])
        //             }else{
        //                 resolve(new Array<ICustomerRecord>())
        //             }
        //         });
                
        //     });
            
        // })
    }
}