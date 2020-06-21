import { Configuration } from "msal";

export const GraphScopes = ['User.Read'];

export const ClientId = "YOURCLIENTID";

export const SharePointUrl = "https://YOURTENANTNAME.sharepoint.com";

export const SharePointScope = [`${SharePointUrl}/.default`];

export const AzureScope = ['https://YOURFUNCTIONURI.azurewebsites.net/user_impersonation'];

export const AzureFunctionUri = "YOURFUNCTIONURL";

export const MsalConfiguration: Configuration = {
  auth: {
    clientId: ClientId,
    authority: "https://login.microsoftonline.com/common",
    redirectUri: "https://localhost:3000/login.html",
    navigateToLoginRequestUrl: false
  },
  cache: {
    cacheLocation: "localStorage", // Needed to avoid "User login is required" error.
    storeAuthStateInCookie: true // Recommended to avoid certain IE/Edge issues.
  }
};