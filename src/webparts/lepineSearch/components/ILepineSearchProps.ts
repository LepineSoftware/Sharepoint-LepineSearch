import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ILepineSearchPreset } from "../models/ISearchResult";

export interface ILepineSearchProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  
  // Add these required fields:
  context: WebPartContext;
  selectedSiteUrl: string;
  selectedLibraryIds: string[];
  presets: ILepineSearchPreset[];
}