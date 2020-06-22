import { loginService } from "../services/services";
import IAuthToken from "../models/IAuthToken";
import { AuthResponse } from "msal/lib-commonjs/AuthResponse";
import * as Msal from "msal";
import * as config from "./configuration";

let loginDialog: Office.Dialog;
let logoutDialog: Office.Dialog;
const url = location.protocol + "//" + location.hostname + (location.port ? ":" + location.port : "");

// This handler responds to the success or failure message that the pop-up dialog receives from the identity provider
// and access token provider.
async function processMessage(arg) {
  let messageFromDialog = JSON.parse(arg.message);

  if (messageFromDialog.status === "success") {
    // We now have a valid access token.
    loginDialog.close();
    const response: AuthResponse = messageFromDialog.result;
    const userAgentApp = new Msal.UserAgentApplication(config.MsalConfiguration);
    const spReqObj = {
      account: response.account,
      scopes: config.SharePointScope
    };
    const azureReqObj = {
      account: response.account,
      scopes: config.AzureScope
    };

    // request SharePoint Access Token
    const spTokenResponse = await userAgentApp.acquireTokenSilent(spReqObj);

    // Request Azure Access Token
    const azureTokenResponse = await userAgentApp.acquireTokenSilent(azureReqObj);

    const tokens: IAuthToken = {
      graphToken: response.accessToken,
      sharePointToken: spTokenResponse.accessToken,
      azureToken: azureTokenResponse.accessToken
    };

    loginService.setTokens(tokens);
  } else {
    // Something went wrong with authentication or the authorization of the web application.
    loginDialog.close();
    loginService.setErrorMessage(JSON.stringify(messageFromDialog.error.toString()));
  }
}

// Use the Office dialog API to open a pop-up and display the sign-in page for the identity provider.
export const loginToOffice365 = () => {
  var fullUrl = url + "/login.html";

  // height and width are percentages of the size of the parent Office application, e.g., PowerPoint, Excel, Word, etc.
  Office.context.ui.displayDialogAsync(fullUrl, { height: 60, width: 30 }, function(result) {
    console.log("Dialog has initialized. Wiring up events");
    loginDialog = result.value;
    loginDialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
  });
};

export const logoutFromO365 = async () => {
  var fullUrl = url + "/logout.html";
  Office.context.ui.displayDialogAsync(fullUrl, { height: 40, width: 30 }, result => {
    if (result.status === Office.AsyncResultStatus.Failed) {
      loginService.setErrorMessage(`${result.error.code} ${result.error.message}`);
    } else {
      logoutDialog = result.value;
      logoutDialog.addEventHandler(Office.EventType.DialogMessageReceived, processLogoutMessage);
    }
  });

  const processLogoutMessage = () => {
    loginService.setTokens(null);
    localStorage.setItem("loggedIn","");
    logoutDialog.close();
  };
};
