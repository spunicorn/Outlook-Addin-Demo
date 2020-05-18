import { IHttpClient } from "../utils/IHttpClient";
import { MsalAuthorization } from "../utils/auth/MsalAuthorization";
import { FetchHttpClient } from "../utils/FetchHttpClient";
import { IProfileInfo } from "../model/dto/IProfileInfo";
import { IEmail } from "../model/dto/IEmail";

export class GraphClient {
    protected HttpClient: IHttpClient;
    public constructor() {
        let authorization = new MsalAuthorization("e728e4c0-a694-4432-8624-64813827cf39",
            "https://localhost:3000/taskpane.html",
            "e370efe0-35be-4e74-979e-b57598b68d4f",
            {
                scopes: ["user.read", "mail.read"] 
            });
        this.HttpClient = new FetchHttpClient(authorization);
    }

    public async getMyProfileInformation(): Promise<IProfileInfo> {
        return this.HttpClient.fetch("https://graph.microsoft.com/v1.0/me", {});
    }
    public async searchMyMailbox(query): Promise<IEmail[]> {
        let temp = await this.HttpClient.fetch<{ value: IEmail[] }>(`https://graph.microsoft.com/v1.0/me/messages?$search="${query}"&$select=id,receivedDateTime,sender,subject,bodyPreview`, {});
        return temp.value;
    }
}