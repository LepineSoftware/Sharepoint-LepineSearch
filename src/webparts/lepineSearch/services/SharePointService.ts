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

  // 2. Get Libraries (Helper for a single site)
  public async getLibraries(siteUrl: string): Promise<{ key: string; text: string }[]> {
    try {
        const web = Web(siteUrl).using(SPFx(this._context));
        const libs = await web.lists.filter("(BaseTemplate eq 101 or BaseTemplate eq 109) and Hidden eq false")();
        return libs.map((l: any) => ({ key: l.Id, text: l.Title }));
    } catch (e) {
        console.error(`Failed to load libraries for ${siteUrl}`, e);
        return [];
    }
  }

  // 3. Main Query: Accepts Composite Keys (SiteUrl::LibId)
  public async getFilesFromLibraries(libraryKeys: string[]): Promise<ILepineSearchResult[]> {
    let allFiles: ILepineSearchResult[] = [];

    // Filter: No Folders (FSObjType eq 0) AND No ASPX pages
    const filterQuery = "FSObjType eq 0 and File_x0020_Type ne 'aspx'";

    // Group keys by Site URL to minimize Web object creation
    // Key Format: "https://site/url::LibraryGUID"
    const libsBySite: Record<string, string[]> = {};

    libraryKeys.forEach(key => {
        // Handle potential legacy keys (just GUID) by ignoring them or handling gracefully
        if (key.indexOf('::') === -1) return;

        const [siteUrl, libId] = key.split('::');
        if (!libsBySite[siteUrl]) libsBySite[siteUrl] = [];
        libsBySite[siteUrl].push(libId);
    });

    // Iterate through each Site
    for (const siteUrl in libsBySite) {
      const libIds = libsBySite[siteUrl];
      const web = Web(siteUrl).using(SPFx(this._context));
      
      for (const libId of libIds) {
        let items: any[] = [];
        
        try {
            const list = web.lists.getById(libId);
            
            // STEP A: Get the "Drive ID"
            const listEntity: any = await list.select("Drive/Id").expand("Drive")();
            const driveId = listEntity.Drive?.Id;

            // STEP B: Get Items
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
                const fileGuid = item.File?.UniqueId; 
                
                // --- THUMBNAIL LOGIC ---
                let thumbUrl = "";
                // Complex types that benefit from Drive API
                const isComplexMedia = ['mp4', 'mov', 'avi', 'wmv', 'mkv', 'heic', 'heif', 'pptx', 'docx', 'xlsx', 'pdf'].indexOf(fileType) > -1;

                if (driveId && fileGuid && isComplexMedia) {
                    // STRATEGY 1: Modern Drive API
                    thumbUrl = `${siteUrl}/_api/v2.1/drives/${driveId}/items/${fileGuid}/thumbnails/0/large/content`;
                } else {
                    // STRATEGY 2: Legacy Fallback
                    thumbUrl = `${siteUrl}/_layouts/15/getpreview.ashx?path=${encodeURIComponent(absUrl)}&resolution=0`;
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
            console.warn(`Library ${libId} query failed in site ${siteUrl}`, error);
        }
      }
    }
    return allFiles;
  }
}