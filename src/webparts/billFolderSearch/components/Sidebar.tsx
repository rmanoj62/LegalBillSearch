import * as React from 'react';
import { useState } from 'react';
import styles from "./BillFolderSearch.module.scss";
import { ChoiceGroup, IChoiceGroupOption, Dropdown, IDropdownOption, DefaultButton, Label } from '@fluentui/react';

export const Sidebar: React.FC = () => {
  const [selectedFederalState, setSelectedFederalState] = useState<string>('federal');
  const [selectedCategory, setSelectedCategory] = useState<string>('all');
  const [selectedCongress, setSelectedCongress] = useState<string | undefined>('119th');

  // Options for filters
  const federalStateOptions: IChoiceGroupOption[] = [
    { key: 'federal', text: 'Federal' },
    { key: 'state', text: 'State' },
    { key: 'both', text: 'Both Federal and State(s)' },
  ];

  const categoryOptions: IChoiceGroupOption[] = [
    { key: 'bills', text: 'Bills' },
    { key: 'proposal', text: 'Proposal' },
    { key: 'all', text: 'All' },
  ];

  const congressOptions: IDropdownOption[] = [
    { key: '119th', text: '119th Congress' },
    { key: '118th', text: '118th Congress' },
    { key: '117th', text: '117th Congress' },
  ];

  return (
    <div className={styles.sidebar}>
      <h3 className={styles.heading}>Filter Results</h3>

      {/* Federal/State Filter */}
      <div className={styles.filterSection}>
        <Label>Federal/State</Label>
        <ChoiceGroup
          selectedKey={selectedFederalState}
          options={federalStateOptions}
          onChange={(ev, option) => setSelectedFederalState(option?.key || 'federal')}
        />
      </div>

      {/* Category Filter */}
      <div className={styles.filterSection}>
        <Label>Category</Label>
        <ChoiceGroup
          selectedKey={selectedCategory}
          options={categoryOptions}
          onChange={(ev, option) => setSelectedCategory(option?.key || 'all')}
        />
      </div>

      {/* Congress Filter */}
      <div className={styles.filterSection}>
        <Label>Congress</Label>
        <Dropdown
          placeholder="Select Congress"
          options={congressOptions}
          selectedKey={selectedCongress}
          onChange={(ev, option) => setSelectedCongress(option?.key?.toString())}
        />
      </div>

      {/* Apply Filters Button */}
      <div className={styles.actionButton}>
        <DefaultButton text="Apply Filters" onClick={() => console.log('Filters applied')} />
      </div>
    </div>
  );
};

export default Sidebar;
