import * as React from 'react';
import { SearchBox, ISearchBoxStyles } from '@fluentui/react/lib/SearchBox';
import { Stack, Text } from '@fluentui/react';

interface ISearchBarProps {
  onSearch: (query: string) => void;
  currentValue: string; // Passed from parent (debounced value usually, but here used for initial)
}

// Internal state to handle immediate UI updates while parent debounces
interface ISearchBarState {
  localValue: string;
}

const searchBoxStyles: Partial<ISearchBoxStyles> = { root: { width: '100%' } };

export default class LepineSearchResultsSearchBar extends React.Component<ISearchBarProps, ISearchBarState> {
  
  constructor(props: ISearchBarProps) {
    super(props);
    this.state = {
      localValue: props.currentValue || ''
    };
  }

  // Sync state if parent updates query via presets or other means
  public componentDidUpdate(prevProps: ISearchBarProps) {
    if (prevProps.currentValue !== this.props.currentValue) {
      this.setState({ localValue: this.props.currentValue });
    }
  }

  private _onChange = (_: any, newValue?: string) => {
    const val = newValue || '';
    this.setState({ localValue: val });
    // This triggers the parent's handler (which we will debounce in the parent)
    this.props.onSearch(val);
  }

  public render(): React.ReactElement<ISearchBarProps> {
    return (
      <Stack tokens={{ childrenGap: 5 }}>
        <Text variant="medium" styles={{ root: { fontWeight: '600' } }}>
          Search Documents
        </Text>
        <SearchBox
          placeholder="Search by filename..."
          onSearch={(newValue) => this.props.onSearch(newValue)}
          onClear={() => {
            this.setState({ localValue: '' });
            this.props.onSearch('');
          }}
          onChange={this._onChange}
          value={this.state.localValue}
          styles={searchBoxStyles}
        />
      </Stack>
    );
  }
}