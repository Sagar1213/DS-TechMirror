import { IPickerTerms } from "@pnp/spfx-controls-react/lib/TaxonomyPicker";

export interface ICustomFormState {
  email: string;
  mobile: string;
  address: string;
  availableOnWeekdays: boolean;
  managerApproval: string;
  empPeoplePicker: any[];
  Courses: IPickerTerms;
  multiCourses: IPickerTerms;
  hideDialog: boolean;
  defaultEmp: string[];
}
