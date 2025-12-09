import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import { FluentProvider, webLightTheme } from '@fluentui/react-components';

// PnP Controls
import { PropertyFieldMultiSelect } from '@pnp/spfx-property-controls/lib/PropertyFieldMultiSelect';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

import LepineSearch from './components/LepineSearch';
import { ILepineSearchProps } from './components/ILepineSearchProps';
import { SharePointService } from './services/SharePointService';
import { ILepineSearchPreset } from './models/ISearchResult';

export interface ILepineSearchWebPartProps {
  selectedSiteUrls: string[]; 
  selectedLibraryIds: string[];
  presets: ILepineSearchPreset[];
}

export default class LepineSearchWebPart extends BaseClientSideWebPart<ILepineSearchWebPartProps> {

  private _service: SharePointService;
  private _siteOptions: any[] = [];
  private _libraryOptions: any[] = [];

  public render(): void {
    const element: React.ReactElement<ILepineSearchProps> = React.createElement(
      LepineSearch,
      {
        description: this.properties.presets ? "Search Configured" : "Please configure",
        isDarkTheme: !!this.context.sdks.microsoftTeams, 
        environmentMessage: "",
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context,
        // We now pass ONLY the composite keys. The component doesn't need the site URL explicitly anymore.
        selectedLibraryIds: this.properties.selectedLibraryIds || [],
        presets: this.properties.presets || []
      }
    );

      const wrappedElement = React.createElement(
        FluentProvider,
        { theme: webLightTheme },
        element
      );

    ReactDom.render(wrappedElement, this.domElement);
  }

  protected async onInit(): Promise<void> {
    await super.onInit();
    this._service = new SharePointService(this.context);
    
    // 1. Load available sites
    try {
        const sites = await this._service.getAllSites();
        this._siteOptions = sites.map(s => ({ key: s.key, text: s.text }));
    } catch (e) {
        console.error("Error loading sites", e);
    }

    // 2. Pre-load library options if sites are already selected (Handles page refresh)
    if (this.properties.selectedSiteUrls && this.properties.selectedSiteUrls.length > 0) {
       await this._loadLibrariesForSites(this.properties.selectedSiteUrls);
    }
  }

  // Refactored helper to load libraries from MULTIPLE sites
  private async _loadLibrariesForSites(siteUrls: string[]): Promise<void> {
      this._libraryOptions = [];
      const promises: Promise<any[]>[] = [];

      // Create a promise for each selected site
      siteUrls.forEach(url => {
          promises.push(
              this._service.getLibraries(url)
                .then(libs => {
                    // Find site title for display purposes (Optional, but nice UI)
                    const siteTitle = this._siteOptions.find(opt => opt.key === url)?.text || "Site";
                    
                    // Map to Composite Key: "SiteURL::LibraryID"
                    return libs.map(lib => ({
                        key: `${url}::${lib.key}`,
                        text: `${siteTitle} > ${lib.text}` 
                    }));
                })
                .catch(err => {
                    console.error(`Error loading libs for ${url}`, err);
                    return [];
                })
          );
      });

      // Wait for all sites to respond
      const results = await Promise.all(promises);
      
      // Flatten the results array
      this._libraryOptions = results.reduce((acc, val) => acc.concat(val), []);
      
      // Refresh the pane to show new options
      this.context.propertyPane.refresh();
  }

  protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): Promise<void> {
    
    if (propertyPath === 'selectedSiteUrls') {
      // NEW: Clear cache to ensure fresh data for new site selection
      this._service.clearCache();

      // If sites change, reload libraries for ALL selected sites
      await this._loadLibrariesForSites(newValue as string[] || []);
      
      // IMPORTANT: Removed the line that reset selectedLibraryIds
      // This ensures previous selections persist if their key (Site::Lib) is still valid
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: "Lepine Search Configuration" },
          groups: [
            {
              groupName: "Data Sources",
              groupFields: [
                PropertyFieldMultiSelect('selectedSiteUrls', {
                  key: 'multiSelectSites',
                  label: "Select Sites",
                  options: this._siteOptions,
                  selectedKeys: this.properties.selectedSiteUrls
                }),
                PropertyFieldMultiSelect('selectedLibraryIds', {
                  key: 'multiSelectLibs',
                  label: "Select Document Libraries",
                  options: this._libraryOptions,
                  selectedKeys: this.properties.selectedLibraryIds,
                  // Only disable if NO sites are selected
                  disabled: !this.properties.selectedSiteUrls || this.properties.selectedSiteUrls.length === 0
                })
              ]
            },
            {
              groupName: "Search Presets",
              groupFields: [
                PropertyFieldCollectionData('presets', {
                  key: 'collectionData',
                  label: 'Manage Search Presets',
                  panelHeader: 'Manage Presets',
                  manageBtnLabel: 'Manage Presets',
                  value: this.properties.presets,
                  fields: [
                    {
                      id: 'name',
                      title: 'Button Label',
                      type: CustomCollectionFieldType.string,
                      required: true
                    },
                    {
                      id: 'query',
                      title: 'Search Term',
                      type: CustomCollectionFieldType.string,
                      required: true
                    }
                  ]
                })
              ]
            }
          ]
        }
      ]
    };
  }
}