import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFI, SPFx } from "@pnp/sp";
import { Web } from "@pnp/sp/webs"; // Import Web specifically
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/search";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import { ILepineSearchResult } from "../models/ISearchResult";

export class SharePointService {
  private _sp: SPFI;
  private _context: WebPartContext;

  constructor(context: WebPartContext) {
    this._context = context;
    this._sp = spfi().using(SPFx(context));
  }

  // 1. Get All Sites
  public async getAllSites(): Promise<{ key: string; text: string }[]> {
    const results = await this._sp.search({
      Querytext: "contentclass:STS_Site contentclass:STS_Web",
      SelectProperties: ["Title", "Path", "SiteId"],
      RowLimit: 500,
      TrimDuplicates: false
    });

    // Fix TS2322: Handle undefined values with logical OR (|| "")
    return results.PrimarySearchResults.map((r: any) => ({ 
      key: r.Path || "", 
      text: r.Title || "" 
    }));
  }

  // 2. Get Libraries for a specific site URL
  public async getLibraries(siteUrl: string): Promise<{ key: string; text: string }[]> {
    // Fix TS2339: Use 'Web(url)' factory instead of 'spfi(url)'
    // We must pass the SPFx context again for the new web instance to handle auth
    const web = Web(siteUrl).using(SPFx(this._context));
    
    const libs = await web.lists.filter("BaseTemplate eq 101 and Hidden eq false")();
    
    // Fix TS7006: Explicitly type 'l' as any
    return libs.map((l: any) => ({ key: l.Id, text: l.Title }));
  }

  // 3. Main Query: Get Files from selected libraries
  public async getFilesFromLibraries(siteUrl: string, libraryIds: string[]): Promise<ILepineSearchResult[]> {
    const web = Web(siteUrl).using(SPFx(this._context)); 
    
    let allFiles: ILepineSearchResult[] = [];

    for (const libId of libraryIds) {
      try {
        // UPDATED QUERY: Removed "Keywords" to prevent 400 Error
        const items = await web.lists.getById(libId).items
          .select("Id", "FileLeafRef", "FileRef", "EncodedAbsUrl", "File_x0020_Type", "TaxCatchAll/Term")
          .expand("File", "TaxCatchAll")();

        const mapped: ILepineSearchResult[] = items.map((item: any) => ({
          id: item.Id.toString(),
          name: item.FileLeafRef,
          location: siteUrl, 
          fileType: item.File_x0020_Type,
          href: item.EncodedAbsUrl,
          thumbnailUrl: `${siteUrl}/_layouts/15/getpreview.ashx?path=${encodeURIComponent(item.EncodedAbsUrl)}`, 
          // Map tags from TaxCatchAll (Managed Metadata)
          tags: item.TaxCatchAll ? item.TaxCatchAll.map((t: any) => t.Term) : [],
          metadata: {},
          parentLibraryId: libId,
          parentSiteUrl: siteUrl
        }));
        allFiles = [...allFiles, ...mapped];

      } catch (error) {
        console.warn(`Failed to load items from library ${libId}`, error);
        // Continue loop so one broken library doesn't stop the whole web part
      }
    }
    return allFiles;
  }
}