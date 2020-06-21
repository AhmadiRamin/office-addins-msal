import { Client } from "@microsoft/microsoft-graph-client";
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";

export default class GraphController {
  private client: Client;
  constructor(token: string) {
    const options = {
      defaultVersion: "beta",
      debugLogging: true,      
      authProvider: done => {
        done(null, token);
      }
    };
    this.client = Client.init(options);
  }

  public getClient() {
    return this.client;
  }

  public async getUserInformation() {
    try {
      const userDetails: MicrosoftGraph.User = await this.client.api("/me").get();
      return userDetails;
    } catch (error) {
      throw error;
    }
  }

  public async getUserPhoto(){    
    const photo = await this.client.api("/me/photo/$value").options({encoding:null}).get();
    return Buffer.from(photo).toString('base64');
    
  }
}
