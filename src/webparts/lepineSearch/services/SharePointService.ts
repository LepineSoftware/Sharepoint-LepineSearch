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

// CONFIGURATION: Add your specific Managed Metadata Column Internal Names here
const TARGET_FIELDS = ['Department', 'Project', 'DocType']; 

export class SharePointService {
  private _sp: SPFI;
  private _context: WebPartContext;

  constructor(context: WebPartContext) {
    this._context = context;
    this._sp = spfi().using(SPFx(context));
  }

  public clearCache(): void {
    try {
        sessionStorage.clear();
    } catch (e) {
        console.warn("Failed to clear session storage cache", e);
    }
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
    
    // Standard fields usually available
    const standardSelect = ["Id", "UniqueId", "FileLeafRef", "FileRef", "EncodedAbsUrl", "File_x0020_Type", "File/UniqueId", "TaxCatchAll/Term", "Modified", "Editor/Title", "File/Length"];
    const standardExpand = ["File", "TaxCatchAll", "Editor"];

    // Combine with target fields
    const enhancedSelect = [...standardSelect, ...TARGET_FIELDS];
    const enhancedExpand = [...standardExpand, ...TARGET_FIELDS];

    let items: any[] = [];
    let driveId: string | undefined;
    
    const list = web.lists.getById(libId);

    try {
        const [listEntity, fetchedItems] = await Promise.all([
            list.select("Id", "Drive/Id")
                .expand("Drive")
                .using(Caching({ store: "session" }))()
                .catch((e: any) => {
                    console.warn(`[LepineSearch] Could not fetch Drive ID for list ${libId}.`, e);
                    return {}; 
                }),
            
            list.items
                .filter(filterQuery)
                .select(...enhancedSelect)
                .expand(...enhancedExpand)
                .top(5000)
                .using(Caching({ store: "session" }))()
        ]);

        items = fetchedItems;
        if (listEntity && listEntity.Drive) {
            driveId = listEntity.Drive.Id;
        }
        
    } catch (enhancedError) {
        console.warn(`[LepineSearch] Custom columns failed. Reverting to standard search. Error:`, enhancedError);
        
        try {
            // RETRY: Fetch standard items + Drive ID
            const [listEntity, fetchedItems] = await Promise.all([
                list.select("Id", "Drive/Id")
                    .expand("Drive")
                    .using(Caching({ store: "session" }))()
                    .catch((e: any) => ({})),
                
                list.items
                    .filter(filterQuery)
                    .select(...standardSelect)
                    .expand(...standardExpand)
                    .top(5000)
                    .using(Caching({ store: "session" }))()
            ]);

            items = fetchedItems;
             if (listEntity && listEntity.Drive) {
                driveId = listEntity.Drive.Id;
            }
        } catch (fallbackError) {
             console.error(`Library ${libId} completely failed query`, fallbackError);
             return [];
        }
    }

    try {
        return items.map((item: any) => {
            let fileType = item.File_x0020_Type;
            if (!fileType && item.FileLeafRef) {
                const parts = item.FileLeafRef.split('.');
                if (parts.length > 1) fileType = parts.pop();
            }
            fileType = (fileType || "").toLowerCase();

            const absUrl = item.EncodedAbsUrl || `${siteUrl}${item.FileRef}`;
            const fileGuid = item.File?.UniqueId; 
            
            const isVideo = ['mp4', 'mov', 'avi', 'wmv', 'mkv', 'webm'].indexOf(fileType) > -1;
            
            let thumbUrl = "";

            // --- IMPROVED THUMBNAIL LOGIC ---
            if (driveId && fileGuid) {
                if (isVideo) {
                    thumbUrl = `${siteUrl}/_api/v2.1/drives/${driveId}/items/${fileGuid}/thumbnails/0/c1920x1080/content?prefer=noRedirect,closestavailablesize,extendCacheMaxAge`;
                } else {
                    thumbUrl = `${siteUrl}/_api/v2.1/drives/${driveId}/items/${fileGuid}/thumbnails/0/large/content`;
                }
            } else {
                if (isVideo && fileGuid) {
                    // Refactored fallback for videos to use VideoService
                    thumbUrl = `${siteUrl}/_api/VideoService/Channels('${libId}')/Videos('${fileGuid}')/ThumbnailStream`;
                } else {
                    // Fallback using legacy GetPreview
                    thumbUrl = `${siteUrl}/_layouts/15/getpreview.ashx?path=${encodeURIComponent(absUrl)}&resolution=0`;
                }
            }

            const groups: Record<string, string[]> = {};
            
            TARGET_FIELDS.forEach(field => {
                const val = item[field];
                if (val) {
                    if (Array.isArray(val)) {
                        groups[field] = val.map(v => {
                            if (typeof v === 'object') return v.Label || v.Term || v.Title || "Unknown";
                            return String(v);
                        });
                    } 
                    else if (typeof val === 'object' && (val.Label || val.Term || val.Title)) {
                        groups[field] = [val.Label || val.Term || val.Title];
                    } 
                    else {
                         groups[field] = [String(val)];
                    }
                }
            });

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
                groupedTags: groups,
                modified: item.Modified,
                modifiedBy: item.Editor?.Title,
                fileSize: item.File?.Length ? parseInt(item.File.Length) : 0,
                metadata: {},
                parentLibraryId: libId,
                parentSiteUrl: siteUrl
            };
        });

    } catch (error) {
        console.warn(`Error processing items for library ${libId}`, error);
        return [];
    }
  }
}