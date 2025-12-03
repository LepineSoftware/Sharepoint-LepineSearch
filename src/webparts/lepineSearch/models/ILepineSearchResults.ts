export interface ILepineSearchResult {
  id: string;
  name: string;
  location: string; // Site Name or Library Name
  fileType: string;
  thumbnailUrl: string;
  href: string;
  tags: string[]; // Enterprise keywords / managed metadata
  metadata: Record<string, any>; // Flexible object for other columns
  parentLibraryId: string;
  parentSiteUrl: string;
}

export interface ILepineSearchPreset {
  name: string;
  query: string;
}