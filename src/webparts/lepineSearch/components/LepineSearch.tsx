import * as React from 'react';
import { ILepineSearchProps } from './ILepineSearchProps';
import { Stack } from '@fluentui/react'; // Removed unused MessageBar imports
import { SharePointService } from '../services/SharePointService';
import { ILepineSearchResult } from '../models/ISearchResult';
import LepineSearchResultsSearchBar from './LepineSearchResultsSearchBar';
import LepineSearchResultsFilters from './LepineSearchResultsFilters';
import LepineSearchResultsContent from './LepineSearchResultsContent';
import LepineSearchResultsPresets from './LepineSearchResultsPresets';

export interface ILepineSearchState {
  allItems: ILepineSearchResult[];
  filteredItems: ILepineSearchResult[];
  availableTags: string[];
  isLoading: boolean;
  searchQuery: string;
  selectedFilters: string[];
}

export default class LepineSearch extends React.Component<ILepineSearchProps, ILepineSearchState> {
  private _spService: SharePointService;

  constructor(props: ILepineSearchProps) {
    super(props);
    this.state = {
      allItems: [],
      filteredItems: [],
      availableTags: [],
      isLoading: true,
      searchQuery: '',
      selectedFilters: []
    };
    this._spService = new SharePointService(this.props.context);
  }

  public async componentDidMount() {
    await this._loadData();
  }

  public async componentDidUpdate(prevProps: ILepineSearchProps) {
    if (prevProps.selectedLibraryIds !== this.props.selectedLibraryIds) {
      await this._loadData();
    }
  }

  private _loadData = async () => {
    this.setState({ isLoading: true });
    
    if(!this.props.selectedSiteUrl || !this.props.selectedLibraryIds) {
        this.setState({ isLoading: false });
        return;
    }

    const items = await this._spService.getFilesFromLibraries(this.props.selectedSiteUrl, this.props.selectedLibraryIds);
    
    // Fix TS2550 & TS7006: Replace flatMap with reduce for compatibility
    // Fix TS2322: Explicitly cast the result
    const allTagsRaw = items.reduce<string[]>((acc, item) => {
        return acc.concat(item.tags || []);
    }, []);

    const uniqueTags = Array.from(new Set(allTagsRaw));

    this.setState({
      allItems: items,
      filteredItems: items,
      availableTags: uniqueTags,
      isLoading: false
    });
  }

  private _handleSearch = (query: string) => {
    this.setState({ searchQuery: query }, this._applyFilters);
  }

  private _handleFilterChange = (selectedTags: string[]) => {
    this.setState({ selectedFilters: selectedTags }, this._applyFilters);
  }

  private _applyFilters = () => {
    const { allItems, searchQuery, selectedFilters } = this.state;
    
    let result = allItems;

    if (searchQuery) {
      result = result.filter(i => i.name.toLowerCase().includes(searchQuery.toLowerCase()));
    }

    if (selectedFilters.length > 0) {
      result = result.filter(i => i.tags.some(tag => selectedFilters.indexOf(tag) > -1));
    }

    this.setState({ filteredItems: result });
  }

  public render(): React.ReactElement<ILepineSearchProps> {
    return (
      <Stack tokens={{ childrenGap: 20 }} style={{ padding: 20 }}>
        
        <LepineSearchResultsPresets 
          presets={this.props.presets || []} 
          onPresetSelected={(query) => this._handleSearch(query)} 
        />

        <LepineSearchResultsSearchBar 
          onSearch={this._handleSearch} 
          currentValue={this.state.searchQuery}
        />

        <Stack horizontal tokens={{ childrenGap: 20 }} styles={{ root: { flexWrap: 'wrap' }}}>
          
          <Stack.Item styles={{ root: { width: '250px' } }}>
            <LepineSearchResultsFilters 
              availableTags={this.state.availableTags}
              onFilterApply={this._handleFilterChange}
            />
          </Stack.Item>

          <Stack.Item grow>
            <LepineSearchResultsContent 
              items={this.state.filteredItems} 
              isLoading={this.state.isLoading}
            />
          </Stack.Item>

        </Stack>
      </Stack>
    );
  }
}