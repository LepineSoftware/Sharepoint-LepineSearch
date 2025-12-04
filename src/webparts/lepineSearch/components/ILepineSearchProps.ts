import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ILepineSearchPreset } from "../models/ISearchResult";

export interface ILepineSearchProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  
  context: WebPartContext;
  // We only need the library IDs (which now contain the site info)
  selectedLibraryIds: string[];
  presets: ILepineSearchPreset[];
}