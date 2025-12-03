import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

// PnP Controls
import { PropertyFieldMultiSelect } from '@pnp/spfx-property-controls/lib/PropertyFieldMultiSelect';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

import LepineSearch from './components/LepineSearch';
import { ILepineSearchProps } from './components/ILepineSearchProps';
import { SharePointService } from './services/SharePointService';
import { ILepineSearchPreset } from './models/ISearchResult';

export interface ILepineSearchWebPartProps {
  selectedSiteUrls: string[]; // Multi-select returns an array
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
        isDarkTheme: !!this.context.sdks.microsoftTeams, // basic check
        environmentMessage: "",
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context,
        // The component expects a single URL string for now, but our property pane allows multi. 
        // We take the first selected site.
        selectedSiteUrl: (this.properties.selectedSiteUrls && this.properties.selectedSiteUrls.length > 0) 
            ? this.properties.selectedSiteUrls[0] 
            : '',
        selectedLibraryIds: this.properties.selectedLibraryIds || [],
        presets: this.properties.presets || []
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    await super.onInit();
    // Initialize our service with the WebPart Context
    this._service = new SharePointService(this.context);
    
    // Load initial sites for the property pane dropdown
    try {
        const sites = await this._service.getAllSites();
        this._siteOptions = sites.map(s => ({ key: s.key, text: s.text }));
    } catch (e) {
        console.error("Error loading sites", e);
    }
  }

  // Handle Logic when Site selection changes to reload libraries
  protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): Promise<void> {
    
    if (propertyPath === 'selectedSiteUrls' && newValue) {
      // If user selected sites, we load libraries for the FIRST selected site 
      if(newValue.length > 0) {
        const primarySite = newValue[0];
        // Clear options to indicate loading/refresh
        this._libraryOptions = [];
        
        try {
            const libs = await this._service.getLibraries(primarySite);
            this._libraryOptions = libs.map(l => ({ key: l.key, text: l.text }));
            
            // Clear previous library selections to avoid ID conflicts
            this.properties.selectedLibraryIds = [];
            
            // Refresh the pane to show new options
            this.context.propertyPane.refresh();
        } catch (error) {
            console.error("Error loading libraries", error);
        }
      }
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
                // 1. Multi-Select for Sites
                PropertyFieldMultiSelect('selectedSiteUrls', {
                  key: 'multiSelectSites',
                  label: "Select Sites",
                  options: this._siteOptions,
                  selectedKeys: this.properties.selectedSiteUrls
                }),
                // 2. Multi-Select for Libraries
                PropertyFieldMultiSelect('selectedLibraryIds', {
                  key: 'multiSelectLibs',
                  label: "Select Document Libraries",
                  options: this._libraryOptions,
                  selectedKeys: this.properties.selectedLibraryIds,
                  disabled: !this.properties.selectedSiteUrls || this.properties.selectedSiteUrls.length === 0
                })
              ]
            },
            {
              groupName: "Search Presets",
              groupFields: [
                // 3. Collection Data for Presets (Allows user to add/remove presets)
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