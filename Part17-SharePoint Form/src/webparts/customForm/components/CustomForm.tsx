import * as React from "react";
import styles from "./CustomForm.module.scss";
import { ICustomFormProps } from "./ICustomFormProps";
import { ICustomFormState } from "./ICustomFormState";
import { escape } from "@microsoft/sp-lodash-subset";

import {
  Label,
  TextField,
  ChoiceGroup,
  Checkbox,
  IChoiceGroupOption,
  Button,
  Dialog,
  DialogFooter,
  DialogType,
  PrimaryButton,
} from "@fluentui/react";
import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import {
  TaxonomyPicker,
  IPickerTerms,
} from "@pnp/spfx-controls-react/lib/TaxonomyPicker";

import { sp } from "@pnp/sp/presets/all";
const dialogContent = {
  type: DialogType.normal,
  Title: "Message",
  subText: "Your request have been submitted successfully",
  closeButtonAriaLabel: "Close",
};
export default class CustomForm extends React.Component<
  ICustomFormProps,
  ICustomFormState,
  {}
> {
  constructor(props: ICustomFormProps) {
    super(props);
    this.getEmail = this.getEmail.bind(this);
    this.getAddress = this.getAddress.bind(this);
    this.getMobile = this.getMobile.bind(this);
    this.getManagerApproval = this.getManagerApproval.bind(this);
    this.SubmitData = this.SubmitData.bind(this);
    this.getCourse = this.getCourse.bind(this);
    this.getMulipleCourses = this.getMulipleCourses.bind(this);
    this.getWeekDayAvailability = this.getWeekDayAvailability.bind(this);

    this.getEmployee = this.getEmployee.bind(this);
    this.Cancel = this.Cancel.bind(this);
    this.state = {
      email: "",
      mobile: "",
      address: "",
      managerApproval: "",
      availableOnWeekdays: false,
      empPeoplePicker: [],
      Courses: [],
      multiCourses: [],
      hideDialog: true,
      defaultEmp: [""],
    };
  }
  public trainingType: IChoiceGroupOption[] = [
    { key: "yes", text: "yes" },
    { key: "No", text: "No" },
  ];

  /*Read Values */

  /**
   * getEmail
ev,value:string   */
  public getEmail(ev, value: string) {
    this.setState({ email: value });
  }

  public getMobile(ev, value: string) {
    this.setState({ mobile: value });
  }
  public getAddress(ev, value: string) {
    this.setState({ address: value });
  }
  public getWeekDayAvailability(ev, value) {
    console.log(value);
    this.setState({ availableOnWeekdays: value });
  }
  /**
   * getManagerApproval
   */
  public getManagerApproval(ev, value: IChoiceGroupOption) {
    console.log(value);
    this.setState({ managerApproval: value.key });
  }

  /**
   * getEmployee
   */
  public getEmployee(items: any[]) {
    console.log(items);
    let ppl: any[] = [];
    let defaultvalue: string[] = [];

    items.map((item) => {
      ppl.push(item.id);
      defaultvalue.push(item.secondaryText);
    });
    this.setState({ empPeoplePicker: ppl });
    this.setState({ defaultEmp: defaultvalue });
  }
  /**
   * getCourse
   */
  public getCourse(selectedTerm: IPickerTerms) {
    this.setState({ Courses: selectedTerm });
  }
  /**
   * getMultipleCourses
   */
  public getMulipleCourses(selectedTerms: IPickerTerms) {
    this.setState({ multiCourses: selectedTerms });
  }
  /**
   * toggleDialog=
 =>  */
  public toggleDialog = (ev) => {
    this.setState({ hideDialog: true });
  };
  /*Read Values end */
  /**
   * SubmitData
   */
  public SubmitData() {
    let validation: boolean = true;
    if (this.state.email == "") {
      validation = false;
      document
        .getElementById("validation_email")
        .setAttribute("style", "display:block !important");
    }

    if (this.state.mobile == "") {
      validation = false;
      document
        .getElementById("validation_mobile")
        .setAttribute("style", "display:block !important");
    }

    if (this.state.empPeoplePicker.length == 0) {
      validation = false;
      document
        .getElementById("validation_employee")
        .setAttribute("style", "display:block !important");
    }
    if (this.state.Courses.length == 0) {
      validation = false;
      document
        .getElementById("validation_course")
        .setAttribute("style", "display:block !important");
    }
    if (this.state.multiCourses.length == 0) {
      validation = false;
      document
        .getElementById("validation_multicourse")
        .setAttribute("style", "display:block !important");
    }

    if (validation) {
      let multiCourseVal: string = "";

      this.state.multiCourses.map((course) => {
        multiCourseVal += `-1;#${course.name}|${course.key};#`;
      });

      sp.web.lists
        .getByTitle("SPFxCustomForm")
        .fields.getByTitle("MultipleCourses_0")
        .get()
        .then((field) => {
          const fieldInternalName = field.InternalName;
          const data = {
            Title: "Custom List Form",
            AvailableOnWeekdays: this.state.availableOnWeekdays,
            EmployeeNameId: { results: this.state.empPeoplePicker },
            Mobile: this.state.mobile,
            Address: this.state.address,
            Email: this.state.email,
            ManagerApproval: this.state.managerApproval,
            Course: {
              __metadata: { type: "SP.Taxonomy.TaxonomyFieldValue" },
              Label: this.state.Courses[0].name,
              TermGuid: this.state.Courses[0].key,
              WssId: -1,
            },
          };
          data[fieldInternalName] = multiCourseVal;
          sp.web.lists
            .getByTitle("SPFxCustomForm")
            .items.add(data)
            .then(() => {
              this.setState({ hideDialog: false });
            });
        });
    }
  }
  /**
   * Cancel
   */
  public Cancel() {
    this.setState({
      defaultEmp: [""],
      email: "",
      mobile: "",
      address: "",
      availableOnWeekdays: false,
      managerApproval: "No",

      Courses: [],
      multiCourses: [],
    });
  }
  public render(): React.ReactElement<ICustomFormProps> {
    return (
      <div className={styles.customForm}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.title}>Training Request Form</div>
            <br></br>
            <div id="custForm">
              <div className={styles.grid}>
                <div className={styles.gridRow}>
                  <div className={styles.smallCol}>
                    <Label className={styles.label}>
                      Employee Name<span className={styles.validation}>*</span>
                    </Label>
                  </div>
                  <div className={styles.largeCol}>
                    <PeoplePicker
                      context={this.props.context}
                      placeholder="Enter Your Name"
                      ensureUser={true}
                      personSelectionLimit={3}
                      groupName={""} // Leave this blank in case you want to filter from all users
                      showtooltip={false}
                      disabled={false}
                      showHiddenInUI={false}
                      resolveDelay={1000}
                      principalTypes={[PrincipalType.User]}
                      selectedItems={this.getEmployee}
                      defaultSelectedUsers={this.state.defaultEmp}
                    ></PeoplePicker>
                    <div id="validation_employee" className="form-validation">
                      <span>You can't leave this blank</span>
                    </div>
                  </div>

                  <div className={styles.smallCol}>
                    <Label className={styles.label}>
                      Email<span className={styles.validation}>*</span>
                    </Label>
                  </div>
                  <div className={styles.largeCol}>
                    <TextField
                      placeholder="Enter your email here"
                      onChange={this.getEmail}
                      value={this.state.email}
                    ></TextField>
                    <div id="validation_email" className="form-validation">
                      <span>You can't leave this blank</span>
                    </div>
                  </div>

                  <div className={styles.smallCol}>
                    <Label className={styles.label}>
                      Mobile<span className={styles.validation}>*</span>
                    </Label>
                  </div>
                  <div className={styles.largeCol}>
                    <TextField
                      type="number"
                      placeholder="Enter your 10 digit mobile no."
                      onChange={this.getMobile}
                      value={this.state.mobile}
                    ></TextField>

                    <div id="validation_mobile" className="form-validation">
                      <span>You can't leave this blank</span>
                    </div>
                  </div>

                  <div className={styles.smallCol}>
                    <Label className={styles.label}>Address</Label>
                  </div>
                  <div className={styles.largeCol}>
                    <TextField
                      multiline={true}
                      onChange={this.getAddress}
                      value={this.state.address}
                    ></TextField>
                  </div>

                  <div className={styles.smallCol}>
                    <Label className={styles.label}>
                      Choose Your Course
                      <span className={styles.validation}>*</span>
                    </Label>
                  </div>
                  <div className={styles.largeCol}>
                    <TaxonomyPicker
                      termsetNameOrID="Skills"
                      panelTitle="Select Term"
                      label=""
                      context={this.props.context}
                      placeholder="Select Course"
                      isTermSetSelectable={false}
                      onChange={this.getCourse}
                      initialValues={this.state.Courses}
                    ></TaxonomyPicker>
                    <div id="validation_course" className="form-validation">
                      <span>You can't leave this blank</span>
                    </div>
                  </div>

                  <div className={styles.smallCol}>
                    <Label className={styles.label}>
                      Choose Multiple Courses
                      <span className={styles.validation}>*</span>
                    </Label>
                  </div>
                  <div className={styles.largeCol}>
                    <TaxonomyPicker
                      allowMultipleSelections={true}
                      termsetNameOrID="Skills"
                      panelTitle="Select Term"
                      label=""
                      context={this.props.context}
                      placeholder="Select Course"
                      isTermSetSelectable={false}
                      onChange={this.getMulipleCourses}
                      initialValues={this.state.multiCourses}
                    ></TaxonomyPicker>
                    <div
                      id="validation_multicourse"
                      className="form-validation"
                    >
                      <span>You can't leave this blank</span>
                    </div>
                  </div>

                  <div className={styles.smallCol}>
                    <Label className={styles.label}>
                      Do you have manager approval
                    </Label>
                  </div>
                  <div className={styles.largeCol}>
                    <ChoiceGroup
                      options={this.trainingType}
                      onChange={this.getManagerApproval}
                      defaultSelectedKey={this.state.managerApproval}
                      selectedKey={this.state.managerApproval}
                    ></ChoiceGroup>{" "}
                  </div>

                  <div className={styles.smallCol}>
                    <Label className={styles.label}>
                      Available On Weekdays
                    </Label>
                  </div>
                  <div className={styles.largeCol}>
                    <Checkbox
                      label="Yes"
                      onChange={this.getWeekDayAvailability}
                      checked={this.state.availableOnWeekdays}
                    ></Checkbox>
                  </div>
                  <Dialog
                    dialogContentProps={dialogContent}
                    hidden={this.state.hideDialog}
                    onDismiss={this.toggleDialog}
                  >
                    <DialogFooter>
                      <PrimaryButton
                        text="Close"
                        onClick={this.toggleDialog}
                      ></PrimaryButton>
                    </DialogFooter>
                  </Dialog>

                  <div className={styles.largeCol}>
                    <Button
                      className={styles.button}
                      text="Submit"
                      onClick={this.SubmitData}
                    ></Button>
                    <Button
                      className={styles.button}
                      text="Cancel"
                      onClick={this.Cancel}
                    ></Button>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
