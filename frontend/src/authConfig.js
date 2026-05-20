import { PublicClientApplication } from "@azure/msal-browser";

export const msalConfig = {
    auth: {
        clientId: "9fb4c731-7cca-4933-a68a-8a76b3b311fa", // User needs to replace this
        authority: "https://login.microsoftonline.com/common",
        redirectUri: window.location.origin,
    },
    cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: false,
    }
};

export const loginRequest = {
    scopes: ["https://management.azure.com/.default"]
};

export const msalInstance = new PublicClientApplication(msalConfig);
