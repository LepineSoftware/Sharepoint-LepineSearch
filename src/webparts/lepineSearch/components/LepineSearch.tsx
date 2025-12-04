import * as React from 'react';
import { ILepineSearchProps } from './ILepineSearchProps';
import { 
  Stack, 
  Toggle, 
  Text, 
  ActionButton, 
  IIconProps, 
  IconButton,
  DefaultButton,
  IButtonStyles,
  Icon
} from '@fluentui/react';
import { SharePointService } from '../services/SharePointService';
import { ILepineSearchResult } from '../models/ISearchResult';
import LepineSearchResultsSearchBar from './LepineSearchResultsSearchBar';
import LepineSearchResultsFilters from './LepineSearchResultsFilters';
import LepineSearchResultsContent from './LepineSearchResultsContent';
import LepineSearchResultsPresets from './LepineSearchResultsPresets';

export interface IFilterGroup {
    category: string;
    values: string[];
}

export interface ILepineSearchState {
  allItems: ILepineSearchResult[];
  filteredItems: ILepineSearchResult[];
  availableFilters: IFilterGroup[]; 
  isLoading: boolean;
  searchQuery: string;
  selectedFilters: string[];
  activeFileKind: string;
  isCardView: boolean;
  isFiltersOpen: boolean; 
}

const filterIcon: IIconProps = { iconName: 'Filter' };
const backIcon: IIconProps = { iconName: 'ChromeBack' };
const cancelIcon: IIconProps = { iconName: 'Cancel' };

const FILE_KINDS: Record<string, string[]> = {
    'Photo': ['jpg', 'jpeg', 'png', 'gif', 'heic', 'heif', 'bmp', 'tiff', 'svg', 'jfif'],
    'Video': ['mp4', 'mov', 'avi', 'wmv', 'mkv', 'webm', 'ogg'],
    'PDF': ['pdf'],
    'Documents': ['doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx', 'txt', 'rtf', 'csv', 'one']
};

export default class LepineSearch extends React.Component<ILepineSearchProps, ILepineSearchState> {
  private _spService: SharePointService;

  constructor(props: ILepineSearchProps) {
    super(props);
    this.state = {
      allItems: [],
      filteredItems: [],
      availableFilters: [],
      isLoading: true,
      searchQuery: '',
      selectedFilters: [],
      activeFileKind: 'All', 
      isCardView: true,
      isFiltersOpen: false
    };
    this._spService = new SharePointService(this.props.context);
  }

  public async componentDidMount() {
    await this._loadData();
  }

  public async componentDidUpdate(prevProps: ILepineSearchProps) {
    // Only reload if library selections change
    if (prevProps.selectedLibraryIds !== this.props.selectedLibraryIds) {
      await this._loadData();
    }
  }

  private _loadData = async () => {
    this.setState({ isLoading: true });
    
    // Check if we have library IDs
    if(!this.props.selectedLibraryIds || this.props.selectedLibraryIds.length === 0) {
        this.setState({ isLoading: false, allItems: [], filteredItems: [] });
        return;
    }

    // Pass the keys directly (they now contain the Site URL info)
    const items = await this._spService.getFilesFromLibraries(this.props.selectedLibraryIds);
    
    const allTags = items.reduce<string[]>((acc, item) => acc.concat(item.tags || []), []);
    const uniqueTags = Array.from(new Set(allTags)).sort();

    const groupedFilters: IFilterGroup[] = [
        { category: "Tags", values: uniqueTags }
    ];

    this.setState({
      allItems: items,
      filteredItems: items,
      availableFilters: groupedFilters,
      isLoading: false
    });
  }

  private _handleSearch = (query: string) => {
    this.setState({ searchQuery: query }, this._applyFilters);
  }

  private _handleFilterChange = (selectedTags: string[]) => {
    this.setState({ 
        selectedFilters: selectedTags,
        isFiltersOpen: false 
    }, this._applyFilters);
  }

  private _handleKindChange = (kind: string) => {
      const newKind = this.state.activeFileKind === kind ? 'All' : kind;
      this.setState({ activeFileKind: newKind }, this._applyFilters);
  }

  private _removeFilter = (tagToRemove: string) => {
      const newFilters = this.state.selectedFilters.filter(t => t !== tagToRemove);
      this._handleFilterChange(newFilters);
  }

  private _applyFilters = () => {
    const { allItems, searchQuery, selectedFilters, activeFileKind } = this.state;
    let result = allItems;

    // 1. Search Query: Check Name OR Tags
    if (searchQuery) {
      const lowerQuery = searchQuery.toLowerCase();
      result = result.filter(i => 
        i.name.toLowerCase().includes(lowerQuery) || 
        (i.tags && i.tags.some(t => t.toLowerCase().includes(lowerQuery)))
      );
    }

    // 2. File Kind Filter
    if (activeFileKind !== 'All') {
        if (activeFileKind === 'Other') {
            const allKnownTypes = [
                ...FILE_KINDS['Photo'], 
                ...FILE_KINDS['Video'], 
                ...FILE_KINDS['PDF'], 
                ...FILE_KINDS['Documents']
            ];
            result = result.filter(i => !allKnownTypes.includes((i.fileType || '').toLowerCase()));
        } else {
            const allowedTypes = FILE_KINDS[activeFileKind];
            result = result.filter(i => allowedTypes.includes((i.fileType || '').toLowerCase()));
        }
    }

    // 3. Explicit Tag Filters (Sidebar)
    if (selectedFilters.length > 0) {
      result = result.filter(i => i.tags && i.tags.some(t => selectedFilters.includes(t)));
    }

    this.setState({ filteredItems: result });
  }

  public render(): React.ReactElement<ILepineSearchProps> {
    const { isFiltersOpen, filteredItems, isCardView, selectedFilters, activeFileKind, isLoading } = this.state;

    const getKindBtnStyles = (kind: string): IButtonStyles => ({
        root: { 
            borderRadius: '20px', 
            border: activeFileKind === kind ? '1px solid #0078d4' : '1px solid #e1e1e1',
            backgroundColor: activeFileKind === kind ? '#e5f3ff' : '#fff',
            height: '32px',
            flexGrow: 1, 
            flexBasis: '1',
            minWidth: 'auto',
            padding: '0 10px'
        },
        label: {
            fontWeight: activeFileKind === kind ? 600 : 400,
            color: activeFileKind === kind ? '#0078d4' : '#323130'
        }
    });

    const showEmptyState = !isLoading && filteredItems.length === 0 && activeFileKind !== 'All';

    return (
      <Stack tokens={{ childrenGap: 20 }} style={{ padding: 20, minHeight: '400px' }}>
        
        {isFiltersOpen ? (
            <Stack tokens={{ childrenGap: 20 }}>
                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }} styles={{root: { borderBottom: '1px solid #eaeaea', paddingBottom: 15 }}}>
                    <IconButton 
                        iconProps={backIcon} 
                        onClick={() => this.setState({ isFiltersOpen: false })}
                        title="Back to Results"
                    />
                    <Text variant="xLarge" styles={{root:{fontWeight:600}}}>Filter Results</Text>
                </Stack>

                <LepineSearchResultsFilters 
                  availableFilters={this.state.availableFilters}
                  activeFilters={selectedFilters}
                  onFilterApply={this._handleFilterChange}
                />
            </Stack>
        ) : (
            <Stack tokens={{ childrenGap: 15 }}>
                
                <LepineSearchResultsPresets 
                    presets={this.props.presets || []} 
                    onPresetSelected={(query) => this._handleSearch(query)} 
                />

                <LepineSearchResultsSearchBar 
                    onSearch={this._handleSearch} 
                    currentValue={this.state.searchQuery}
                />

                <Stack horizontal wrap tokens={{ childrenGap: 10 }} styles={{ root: { width: '100%' } }}>
                    <DefaultButton text="All" onClick={() => this._handleKindChange('All')} styles={getKindBtnStyles('All')} />
                    <DefaultButton text="Photo" onClick={() => this._handleKindChange('Photo')} styles={getKindBtnStyles('Photo')} iconProps={{ iconName: 'Photo2' }} />
                    <DefaultButton text="Video" onClick={() => this._handleKindChange('Video')} styles={getKindBtnStyles('Video')} iconProps={{ iconName: 'Video' }} />
                    <DefaultButton text="PDF" onClick={() => this._handleKindChange('PDF')} styles={getKindBtnStyles('PDF')} iconProps={{ iconName: 'PDF' }} />
                    <DefaultButton text="Documents" onClick={() => this._handleKindChange('Documents')} styles={getKindBtnStyles('Documents')} iconProps={{ iconName: 'WordDocument' }} />
                    <DefaultButton text="Other" onClick={() => this._handleKindChange('Other')} styles={getKindBtnStyles('Other')} />
                </Stack>

                <Stack horizontal horizontalAlign="space-between" verticalAlign="center" styles={{root: { borderBottom: '1px solid #eee', paddingBottom: 10, marginTop: 10}}}>
                    <Text variant="small" styles={{ root: { fontWeight: '600' } }}>
                        Found {filteredItems.length} results
                    </Text>
                    
                    <Stack horizontal tokens={{ childrenGap: 20 }} verticalAlign="center">
                        <ActionButton 
                            iconProps={filterIcon} 
                            allowDisabledFocus 
                            onClick={() => this.setState({ isFiltersOpen: true })}
                        >
                            {selectedFilters.length > 0 
                                ? `Filter Tags (${selectedFilters.length})`
                                : "Filter Tags"}
                        </ActionButton>

                        <Toggle 
                            label="Card View" 
                            inlineLabel 
                            styles={{root: { marginBottom: 0 }}}
                            checked={isCardView}
                            onChange={(ev, checked) => this.setState({ isCardView: !!checked })}
                        />
                    </Stack>
                </Stack>

                {selectedFilters.length > 0 && (
                    <Stack horizontal wrap tokens={{ childrenGap: 10 }}>
                        {selectedFilters.map(tag => (
                            <DefaultButton
                                key={tag}
                                text={tag}
                                iconProps={cancelIcon}
                                onClick={() => this._removeFilter(tag)}
                                styles={{ root: { borderRadius: '20px', height: '32px', padding: '0 15px', border: '1px solid #0078d4', backgroundColor: '#e5f3ff' }, label: { fontWeight: 600, color: '#005a9e' }, icon: { color: '#005a9e', fontSize: 12 } }}
                            />
                        ))}
                        <ActionButton 
                            text="Clear tags" 
                            onClick={() => this._handleFilterChange([])} 
                            styles={{root: {height: '32px', color: '#a80000'}}}
                        />
                    </Stack>
                )}

                {showEmptyState ? (
                    <Stack horizontalAlign="center" tokens={{ childrenGap: 20 }} styles={{ root: { paddingBottom: 40 } }}>
                        <Icon iconName="SearchIssue" styles={{ root: { fontSize: 48, color: '#c8c8c8', marginTop: 100 } }} />
                        <Text variant="large" styles={{ root: { color: '#666' } }}>
                            There are no results for this file kind
                        </Text>
                    </Stack>
                ) : (
                    <LepineSearchResultsContent 
                        items={filteredItems} 
                        isLoading={this.state.isLoading}
                        isCardView={isCardView} 
                    />
                )}
            </Stack>
        )}

      </Stack>
    );
  }
}