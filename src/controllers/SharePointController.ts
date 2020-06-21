import { sp } from "@pnp/sp";
import { Web } from "@pnp/sp/webs";
import { SearchResults } from "@pnp/sp/search";
import "@pnp/sp/lists";
import "@pnp/sp/folders";
import "@pnp/sp/search";
import "@pnp/sp/files";
import * as config from "../helpers/configuration";

const libraryExclusions: string[] = ["Form Templates", "Site Assets", "Style Library", "Teams Wiki Data"];

export default class SharePointController {
  constructor(token: string) {    
    sp.setup({
      sp: {
        baseUrl: `${config.SharePointUrl}`,
        headers: {
          Accept: "application/json;odata=verbose",
          Authorization: `Bearer ${token}`
        }
      }
    });
  }
  
  public async getSiteLibraries(siteUri:string){
    const web = Web(siteUri);
    var response = await web.lists.filter("BaseTemplate eq 101")();    
    var libraries = response.filter(library => libraryExclusions.indexOf(library.Title) === -1);
    return libraries;
  }

  public async getSiteTitle(siteUrl:string){
    const web = Web(siteUrl);
    var response = await web.select("Title")();
    return response.Title;
  }

  public async searchDocuments(query:string) {
    const queryText =`Title:'${query}*'`;    
    let content:SearchResults = await sp.search({
      SelectProperties:["Path","Title","DefaultEncodingURL"],
      Querytext:queryText,
      RowLimit:5
    });
    return content.PrimarySearchResults;
  }

}
