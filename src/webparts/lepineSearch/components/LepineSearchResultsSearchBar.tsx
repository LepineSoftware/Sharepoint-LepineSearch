import * as React from 'react';
import { SearchBox, ISearchBoxStyles } from '@fluentui/react/lib/SearchBox';
import { Stack, Text } from '@fluentui/react';

interface ISearchBarProps {
  onSearch: (query: string) => void;
  currentValue: string;
}

const searchBoxStyles: Partial<ISearchBoxStyles> = { root: { width: '100%' } };

export default class LepineSearchResultsSearchBar extends React.Component<ISearchBarProps, {}> {
  
  public render(): React.ReactElement<ISearchBarProps> {
    return (
      <Stack tokens={{ childrenGap: 5 }}>
        <Text variant="medium" styles={{ root: { fontWeight: '600' } }}>
          Search Documents
        </Text>
        <SearchBox
          placeholder="Search by filename..."
          
          // 1. Handles hitting "Enter" or clicking the search icon
          onSearch={(newValue) => this.props.onSearch(newValue)}
          
          // 2. Handles clicking the "X" to clear
          onClear={() => this.props.onSearch('')}
          
          // 3. CRITICAL: Handles typing or deleting text (Live Filtering)
          onChange={(_, newValue) => this.props.onSearch(newValue || '')}
          
          // Controlled component: value comes from parent state
          value={this.props.currentValue}
          styles={searchBoxStyles}
        />
      </Stack>
    );
  }
}