export interface IDynamicFormProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: any;

  lists: string;
  listName: string;
  onConfigure: () => void;
}
