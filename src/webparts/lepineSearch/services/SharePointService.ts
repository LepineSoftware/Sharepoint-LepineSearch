import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFI, SPFx } from "@pnp/sp";
import { Web } from "@pnp/sp/webs"; 
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

    return results.PrimarySearchResults.map((r: any) => ({ 
      key: r.Path || "", 
      text: r.Title || "" 
    }));
  }

  // 2. Get Libraries
  public async getLibraries(siteUrl: string): Promise<{ key: string; text: string }[]> {
    const web = Web(siteUrl).using(SPFx(this._context));
    const libs = await web.lists.filter("(BaseTemplate eq 101 or BaseTemplate eq 109) and Hidden eq false")();
    return libs.map((l: any) => ({ key: l.Id, text: l.Title }));
  }

  // 3. Main Query: Reliable REST + Correct File GUID for Thumbnails
  public async getFilesFromLibraries(siteUrl: string, libraryIds: string[]): Promise<ILepineSearchResult[]> {
    const web = Web(siteUrl).using(SPFx(this._context)); 
    let allFiles: ILepineSearchResult[] = [];

    // Filter: No Folders (FSObjType eq 0) AND No ASPX pages
    const filterQuery = "FSObjType eq 0 and File_x0020_Type ne 'aspx'";

    for (const libId of libraryIds) {
      let items: any[] = [];
      
      try {
        const list = web.lists.getById(libId);
        
        // STEP A: Get the "Drive ID". Required for the modern thumbnail API.
        const listEntity: any = await list.select("Drive/Id").expand("Drive")();
        const driveId = listEntity.Drive?.Id;

        // STEP B: Get Items. 
        // CRITICAL: We must Select and Expand "File/UniqueId" to get the correct ID for the Drive API.
        try {
            items = await list.items
                .filter(filterQuery)
                .select("Id", "UniqueId", "FileLeafRef", "FileRef", "EncodedAbsUrl", "File_x0020_Type", "File/UniqueId", "TaxCatchAll/Term")
                .expand("File", "TaxCatchAll")();
        } catch (e) {
            console.warn(`Library ${libId} metadata query failed. Retrying basic.`);
            items = await list.items
                .filter(filterQuery)
                .select("Id", "UniqueId", "FileLeafRef", "FileRef", "EncodedAbsUrl", "File_x0020_Type", "File/UniqueId")
                .expand("File")();
        }

        // STEP C: Map Results
        const mapped: ILepineSearchResult[] = items.map((item: any) => {
            let fileType = item.File_x0020_Type;
            if (!fileType && item.FileLeafRef) {
                const parts = item.FileLeafRef.split('.');
                if (parts.length > 1) fileType = parts.pop();
            }
            fileType = (fileType || "").toLowerCase();

            const absUrl = item.EncodedAbsUrl || `${siteUrl}${item.FileRef}`;
            
            // --- THUMBNAIL LOGIC ---
            let thumbUrl = "";
            const isVideo = ['mp4', 'mov', 'avi', 'wmv', 'mkv'].indexOf(fileType) > -1;

            // Use the FILE UniqueId (item.File.UniqueId), not the Item UniqueId
            const fileGuid = item.File?.UniqueId; 

            if (driveId && fileGuid && isVideo) {
                // MODERN VIDEO THUMBNAIL (Drive API)
                // "c400x99999" forces the modern generator to create a cover image
                // We use the Drive ID + File GUID
                thumbUrl = `${siteUrl}/_api/v2.1/drives/${driveId}/items/${fileGuid}/thumbnails/0/c400x99999/content?preferNoRedirect=true`;
            } else {
                // STANDARD IMAGE/DOC THUMBNAIL
                thumbUrl = `${siteUrl}/_layouts/15/getpreview.ashx?path=${absUrl}`;
            }

            return {
                id: item.Id.toString(),
                name: item.FileLeafRef,
                location: siteUrl, 
                fileType: fileType,
                href: absUrl,
                thumbnailUrl: thumbUrl,
                tags: (item.TaxCatchAll && item.TaxCatchAll.length > 0) 
                    ? item.TaxCatchAll.map((t: any) => t.Term) 
                    : [],
                metadata: {},
                parentLibraryId: libId,
                parentSiteUrl: siteUrl
            };
        });
        
        allFiles = [...allFiles, ...mapped];

      } catch (error) {
        console.warn(`Library ${libId} query failed`, error);
      }
    }
    return allFiles;
  }
}