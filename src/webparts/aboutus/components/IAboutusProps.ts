
import { DisplayMode } from '@microsoft/sp-core-library';

export interface IAboutusProps {
  listId: string;
  accordionTitle: string;
  allowZeroExpanded: boolean;
  allowMultipleExpanded: boolean;
  displayMode: DisplayMode;
 
  updateProperty: (value: string) => void;
  onConfigure: () => void;
}
