import { Observable } from "../hooks/observable";
import { loginToOffice365, logoutFromO365 } from "../helpers/msalHelper";
import IAuthToken from "../models/IAuthToken";

export default class LoginService {
  public readonly errorMessage = new Observable<string>("");
  public readonly tokens = new Observable<IAuthToken>(null);

  public setTokens(tokens: IAuthToken) {
    this.tokens.set(tokens);
    sessionStorage.setItem("cachedTokens",JSON.stringify(tokens));
  }

  public setErrorMessage(error: string) {
    this.errorMessage.set(error);
  }

  public async getAccessToken(): Promise<void> {
    const cachedTokens = localStorage.getItem("cachedTokens");
    if (cachedTokens) {
      this.tokens.set(JSON.parse(cachedTokens));
    } else {
      localStorage.setItem("loggedIn", "");
      loginToOffice365();
    }
  }

  public logOut(){
    logoutFromO365();
  }
}
