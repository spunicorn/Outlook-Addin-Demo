import { IHttpClient } from "../utils/IHttpClient";
import { MsalAuthorization } from "../utils/auth/MsalAuthorization";
import { FetchHttpClient } from "../utils/FetchHttpClient";
import { ICustomerRecord } from "../model/dto/ICustomerRecord";
export interface ISharePointNoMetadataResponseICustomer{
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
        // let queryUrl = `https://spunicorntest.sharepoint.com/sites/testteamsite/_api/lists/getByTitle('CustomerInfo')/items?$filter=Email eq '${email}'`;
        let queryUrl = `https://spunicorntest.sharepoint.com/sites/testteamsite/_api/lists/getByTitle('CustomerInfo')/items?$select=Title,DateExpires,DatePurchased,CustomerInfo/Title,CustomerInfo/Id,CustomerInfo/IsVIP,CustomerInfo/Email,ProductInfo/Title,ProductInfo/Id,ProductInfo/IsSupported&$expand=CustomerInfo,ProductInfo&$filter=CustomerInfo/Email eq '${email}'`;
        

        let headers = new Headers({
            'Content-Type': 'application/json',
            'Accept': 'application/json;odata=nometadata'
        })
        let something = await this.HttpClient.fetch<ISharePointNoMetadataResponseICustomer>(queryUrl, { headers: headers });

        console.log(something.value)
        return something.value;
       
    }
}