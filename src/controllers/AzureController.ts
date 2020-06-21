import * as config from "../helpers/configuration";
import axios from "axios";

export default class AzureController {
  constructor(private token: string) {}

  public async callAzureFunction(name:string) {
    const body = {
      name
    };
    
    const response = await axios({
      url: `${config.AzureFunctionUri}`,
      method: "POST",
      data: JSON.stringify(body),
      headers: { Authorization: `Bearer ${this.token}`}
    });
    return response.data;
  }
}
