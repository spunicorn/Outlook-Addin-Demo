import { IAuthorization } from "./IAuthorization";
import * as Msal from "msal";
import { IToken } from "../../model/dto/IToken";

export class MsalAuthorization implements IAuthorization {
    protected msalConfig: Msal.Configuration;
    protected msalInstance;
    protected token: IToken;
    private readonly storageKey: string = "mail-helper-token";
    /**
     * 
     * @param clientId aad app client id: 93dc701d-973c-4fac-ae26-15c54784eb39
     * @param redirectUri redirectUri registered in aad https://localhost:3000
     * @param tenantId id of tenant client is registered in 6bcd5635-bb07-4d76-9e96-0dd01692cbc5
     */

    constructor(protected clientId: string, protected redirectUri: string, tenantId: string, protected loginRequest: {
        scopes: string[];
    }) {
        let resources = new Map<string, string[]>();
        this.msalConfig = {
            auth: {
                clientId: clientId,
                redirectUri: redirectUri,
                authority: "https://login.microsoftonline.com/" + tenantId
            },
            framework: {
                protectedResourceMap: resources
            }
        };
        if (window.location.hash.includes('id_token=')) {
            Office.onReady(() => {
                if (Office.context.ui) {
                    Office.context.ui.messageParent(window.location.hash);
                } else {
                }
            });
        }
        else {
            this.msalInstance = new Msal.UserAgentApplication(this.msalConfig);
            if (sessionStorage.getItem(this.storageKey)) {
                let tempToken = JSON.parse(sessionStorage.getItem(this.storageKey));
                if (new Date(tempToken.expiration) > new Date()) {
                    this.token = tempToken;
                }
            }
            this.msalInstance.handleRedirectCallback(this.handleRedirect.bind(this));
            this.msalInstance.openPopup = () => {
                const dummy = {
                    close() {
                    },
                    location: {
                        assign(url) {
                            Office.context.ui.displayDialogAsync(url, { width: 25, height: 50 }, res => {
                                dummy.close = res.value.close;
                                res.value.addEventHandler(Office.EventType.DialogMessageReceived, ({ message }) =>
                                    dummy.location.href = dummy.location.hash = message
                                );
                            });
                        },
                        href: "",
                        hash: ""
                    }
                };
                return dummy;
            };
            Office.onReady(() => {
                if (Office.context.ui) {
                    Office.context.ui.messageParent(window.location.hash);
                } else {
                }
            });
        }
    }
    private handleRedirect(error, response) {
        if (error) {
            console.log(error)
        };
        console.log(response);
    }
    public async authorize(url: string): Promise<{ key: string; value: string; }> {
        if (!this.token) {
            this.token = await this.authenticateUser(this.loginRequest);
            sessionStorage.setItem(this.storageKey, JSON.stringify(this.token));
        }
        return Promise.resolve({
            key: "Authorization",
            value: "Bearer " + this.token.accessToken
        });
    }



    protected async authenticateUser(loginRequest): Promise<IToken> {
        try {
            await this.msalInstance.loginPopup(loginRequest);
        }
        catch (err) {
            console.log(err);
        }
        if (this.msalInstance.getAccount()) {
            try {
                let token = await this.msalInstance.acquireTokenSilent(loginRequest);
                return token;
            }
            catch (err) {
                try {
                    if (err.name === "InteractionRequiredAuthError") {
                        return this.msalInstance.acquireTokenPopup(loginRequest)
                    }
                    else {
                        throw "Unable to acquire token";
                    }
                }
                catch (innerErr) {
                    throw err;
                }
            }
        }
        else {
            throw "Unable to detect current user";
        }
    }

}