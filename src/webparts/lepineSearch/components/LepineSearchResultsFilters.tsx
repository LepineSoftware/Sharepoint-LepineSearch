import * as React from 'react';
import { Checkbox, PrimaryButton, Stack, Text } from '@fluentui/react';

interface IFilterProps {
  availableTags: string[];
  onFilterApply: (tags: string[]) => void;
}

export default function LepineSearchResultsFilters(props: IFilterProps) {
  const [selected, setSelected] = React.useState<string[]>([]);

  // Fix TS2345: Make 'ev' optional
  const _onChange = (ev?: React.FormEvent<HTMLElement | HTMLInputElement>, isChecked?: boolean, tag?: string) => {
    if (!tag) return;
    
    if (isChecked) {
      setSelected(prev => [...prev, tag]);
    } else {
      setSelected(prev => prev.filter(t => t !== tag));
    }
  };

  return (
    <Stack tokens={{ childrenGap: 10 }} styles={{ root: { borderRight: '1px solid #eaeaea', paddingRight: 10 }}}>
      <Text variant="xLarge">Filters</Text>
      
      <Text variant="mediumPlus" styles={{root:{fontWeight: 'bold'}}}>Tags</Text>
      {props.availableTags.map(tag => (
        <Checkbox 
          key={tag} 
          label={tag} 
          // Pass tag explicitly to the closure
          onChange={(ev, checked) => _onChange(ev, checked, tag)} 
        />
      ))}

      <PrimaryButton 
        text="Apply Filters" 
        onClick={() => props.onFilterApply(selected)} 
        style={{marginTop: 20}}
      />
    </Stack>
  );
}