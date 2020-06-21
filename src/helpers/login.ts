
import * as Msal from "msal";
import * as config from "./configuration";
import { AuthenticationParameters } from "msal";

Office.initialize = function() {
  if (Office.context.ui.messageParent) {
    userAgentApp.handleRedirectCallback(authCallback);
    
    if (localStorage.getItem("loggedIn") === "yes") {
        userAgentApp.acquireTokenRedirect(requestObj);        
    } else {
      userAgentApp.loginRedirect(requestObj);
    }
  }
};

var requestObj: AuthenticationParameters = {
  scopes: config.GraphScopes  
};

let userAgentApp: Msal.UserAgentApplication = new Msal.UserAgentApplication(
  config.MsalConfiguration
);

//const options = new MSALAuthenticationProviderOptions(scopes);
//const authProvider = new ImplicitMSALAuthenticationProvider(userAgentApp, options);

function authCallback(error, response) {
  if (error) {
    Office.context.ui.messageParent(JSON.stringify({ status: "failure", result: error }));
  } else {
    if (response.tokenType === "id_token") {
      
      localStorage.setItem("loggedIn", "yes");
    } else {
      const result: Msal.AuthResponse = response;
      // silently acquire sharepoint access token
      Office.context.ui.messageParent(JSON.stringify({ status: "success", result}));
    }
  }
}
