import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFI, SPFx } from "@pnp/sp";
import { Caching } from "@pnp/queryable";
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

  public async getLibraries(siteUrl: string): Promise<{ key: string; text: string }[]> {
    try {
        const web = Web(siteUrl).using(SPFx(this._context));
        const libs = await web.lists
            .filter("(BaseTemplate eq 101 or BaseTemplate eq 109) and Hidden eq false")
            .using(Caching({ store: "session" }))();
            
        return libs.map((l: any) => ({ key: l.Id, text: l.Title }));
    } catch (e) {
        console.error(`Failed to load libraries for ${siteUrl}`, e);
        return [];
    }
  }

  public async getFilesFromLibraries(libraryKeys: string[]): Promise<ILepineSearchResult[]> {
    const filterQuery = "FSObjType eq 0 and File_x0020_Type ne 'aspx'";
    const libsBySite: Record<string, string[]> = {};

    libraryKeys.forEach(key => {
        if (key.indexOf('::') === -1) return;
        const [siteUrl, libId] = key.split('::');
        if (!libsBySite[siteUrl]) libsBySite[siteUrl] = [];
        libsBySite[siteUrl].push(libId);
    });

    const tasks: Promise<ILepineSearchResult[]>[] = [];

    for (const siteUrl in libsBySite) {
      const libIds = libsBySite[siteUrl];
      const web = Web(siteUrl).using(SPFx(this._context));
      
      for (const libId of libIds) {
        tasks.push(this._fetchSingleLibrary(web, siteUrl, libId, filterQuery));
      }
    }

    const results = await Promise.all(tasks);
    return results.reduce((acc, val) => acc.concat(val), []);
  }

  private async _fetchSingleLibrary(web: any, siteUrl: string, libId: string, filterQuery: string): Promise<ILepineSearchResult[]> {
    try {
        const list = web.lists.getById(libId);

        const [listEntity, items] = await Promise.all([
            list.select("Drive/Id").expand("Drive").using(Caching({ store: "session" }))().catch(() => ({})),
            
            list.items
                .filter(filterQuery)
                // FIX: Changed File_x0020_Size to File/Length
                .select("Id", "UniqueId", "FileLeafRef", "FileRef", "EncodedAbsUrl", "File_x0020_Type", "File/UniqueId", "TaxCatchAll/Term", "Modified", "Editor/Title", "File/Length")
                .expand("File", "TaxCatchAll", "Editor")
                .top(5000)
                .using(Caching({ store: "session" }))()
        ]);

        const driveId = listEntity.Drive?.Id;

        return items.map((item: any) => {
            let fileType = item.File_x0020_Type;
            if (!fileType && item.FileLeafRef) {
                const parts = item.FileLeafRef.split('.');
                if (parts.length > 1) fileType = parts.pop();
            }
            fileType = (fileType || "").toLowerCase();

            const absUrl = item.EncodedAbsUrl || `${siteUrl}${item.FileRef}`;
            const fileGuid = item.File?.UniqueId; 
            
            // Define video types
            const isVideo = ['mp4', 'mov', 'avi', 'wmv', 'mkv', 'webm'].indexOf(fileType) > -1;
            
            let thumbUrl = "";

            if (driveId && fileGuid) {
                // STRATEGY 1: Modern Drive API (Best for everything if Drive ID exists)
                thumbUrl = `${siteUrl}/_api/v2.1/drives/${driveId}/items/${fileGuid}/thumbnails/0/large/content`;
            } else {
                // STRATEGY 2: Fallback Logic
                if (isVideo && fileGuid) {
                    // VIDEO-ONLY FIX: Use guidFile + clientType=modern to avoid 501 error
                    thumbUrl = `${siteUrl}/_layouts/15/getpreview.ashx?guidFile=${fileGuid}&resolution=0&clientType=modern`;
                } else {
                    // STANDARD FALLBACK: Use 'path' for images/docs (keeps them working)
                    thumbUrl = `${siteUrl}/_layouts/15/getpreview.ashx?path=${encodeURIComponent(absUrl)}&resolution=0`;
                }
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
                
                // Mappings
                modified: item.Modified,
                modifiedBy: item.Editor?.Title,
                // FIX: Use File.Length for size
                fileSize: item.File?.Length ? parseInt(item.File.Length) : 0,
                
                metadata: {},
                parentLibraryId: libId,
                parentSiteUrl: siteUrl
            };
        });

    } catch (error) {
        console.warn(`Library ${libId} query failed in site ${siteUrl}`, error);
        return [];
    }
  }
}