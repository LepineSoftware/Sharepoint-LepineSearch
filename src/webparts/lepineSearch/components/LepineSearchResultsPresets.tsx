import * as React from 'react';
import { Stack, DefaultButton, Text, IButtonStyles } from '@fluentui/react';
import { ILepineSearchPreset } from '../models/ISearchResult';

interface IPresetsProps {
  presets: ILepineSearchPreset[];
  onPresetSelected: (query: string) => void;
}

const buttonStyles: IButtonStyles = {
  root: { borderRadius: '20px', padding: '0 20px' } 
};

// Fix TS2322: Change return type to allow null
export default class LepineSearchResultsPresets extends React.Component<IPresetsProps, {}> {
  public render(): React.ReactElement<IPresetsProps> | null {
    const { presets } = this.props;

    if (!presets || presets.length === 0) {
      return null;
    }

    return (
      <Stack tokens={{ childrenGap: 10 }}>
        <Text variant="smallPlus" styles={{ root: { color: '#666' } }}>
          Quick Filters:
        </Text>
        <Stack horizontal wrap tokens={{ childrenGap: 10 }}>
          {presets.map((preset, index) => (
            <DefaultButton
              key={index}
              text={preset.name}
              onClick={() => this.props.onPresetSelected(preset.query)}
              styles={buttonStyles}
              iconProps={{ iconName: 'Search' }}
            />
          ))}
          <DefaultButton
            text="Clear All"
            onClick={() => this.props.onPresetSelected('')}
            styles={buttonStyles}
            iconProps={{ iconName: 'Cancel' }}
          />
        </Stack>
      </Stack>
    );
  }
}