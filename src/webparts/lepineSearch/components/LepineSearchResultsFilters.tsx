import * as React from 'react';
import { 
  makeStyles, 
  shorthands, 
  Checkbox, 
  Button, 
  Text,
  tokens
} from '@fluentui/react-components';
import { 
  FilterRegular, 
  DismissRegular,
  ChevronDownRegular,
  ChevronRightRegular 
} from '@fluentui/react-icons';

// Interface definitions
interface IFilterGroup {
    category: string;
    values: string[];
}

interface IFilterProps {
  availableFilters: IFilterGroup[];
  activeFilters: string[]; // NEW PROP: Received from parent
  onFilterApply: (tags: string[]) => void;
}

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap('15px'),
  },
  groupContainer: {
    display: 'flex',
    flexDirection: 'column',
    marginBottom: '10px',
  },
  groupHeader: {
    display: 'flex',
    alignItems: 'center',
    cursor: 'pointer',
    ...shorthands.padding('5px', '0'),
    ':hover': {
        color: tokens.colorBrandBackground
    }
  },
  checkboxGrid: {
    display: 'flex',
    flexWrap: 'wrap',
    ...shorthands.gap('15px'),
    paddingLeft: '10px',
    marginTop: '5px'
  },
  buttons: {
    display: 'flex',
    ...shorthands.gap('10px'),
    marginTop: '10px'
  }
});

export default function LepineSearchResultsFilters(props: IFilterProps) {
  const styles = useStyles();
  
  // Initialize state using the prop so previous selections are remembered
  const [selected, setSelected] = React.useState<string[]>(props.activeFilters || []);
  
  const [collapsedGroups, setCollapsedGroups] = React.useState<Record<string, boolean>>({});

  const toggleGroup = (category: string) => {
      setCollapsedGroups(prev => ({
          ...prev,
          [category]: !prev[category]
      }));
  };

  const _onChange = (isChecked: boolean, tag: string) => {
    let newSelection = [...selected];
    if (isChecked) {
      newSelection.push(tag);
    } else {
      newSelection = newSelection.filter(t => t !== tag);
    }
    setSelected(newSelection);
  };

  const isChecked = (val: string) => selected.includes(val);

  return (
    <div className={styles.container}>
      <Text size={500} weight="semibold">Select Filters</Text>
      
      {props.availableFilters.map((group) => {
          if(group.values.length === 0) return null;
          const isCollapsed = !!collapsedGroups[group.category];

          return (
            <div key={group.category} className={styles.groupContainer}>
                {/* Header */}
                <div 
                    className={styles.groupHeader} 
                    onClick={() => toggleGroup(group.category)}
                >
                    {isCollapsed ? <ChevronRightRegular /> : <ChevronDownRegular />}
                    <Text weight="semibold" style={{ marginLeft: 5 }}>
                        {group.category}
                    </Text>
                </div>

                {/* Content */}
                {!isCollapsed && (
                    <div className={styles.checkboxGrid}>
                        {group.values.map(val => (
                            <Checkbox 
                                key={val} 
                                label={val} 
                                checked={isChecked(val)}
                                onChange={(ev, data) => _onChange(data.checked as boolean, val)} 
                            />
                        ))}
                    </div>
                )}
            </div>
          );
      })}

      <div className={styles.buttons}>
        <Button 
            appearance="primary"
            icon={<FilterRegular />}
            onClick={() => props.onFilterApply(selected)} 
        >
            Apply ({selected.length})
        </Button>
        <Button 
            appearance="subtle"
            icon={<DismissRegular />}
            onClick={() => { setSelected([]); props.onFilterApply([]); }} 
        >
            Clear All
        </Button>
      </div>
    </div>
  );
}