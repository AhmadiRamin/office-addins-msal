import { ClientId } from './configuration';
import * as msal from 'msal';

(() => {
  // The initialize function must be run each time a new page is loaded
  Office.initialize = () => {

    const config: msal.Configuration = {
      auth: {
          clientId: ClientId,
          redirectUri: 'https://localhost:3000/logoutcomplete.html', 
          postLogoutRedirectUri: 'https://localhost:3000/logoutcomplete.html'
      }
    };

    const userAgentApplication = new msal.UserAgentApplication(config);
    userAgentApplication.logout();
  };
})();
