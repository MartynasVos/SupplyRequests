import * as React from "react";
import { Toggle } from "@fluentui/react/lib/Toggle";
import styles from "./ListWebPart.module.scss";
import { IFiltersProps } from "./ListItems";
import {
  ComboBox,
  DatePicker,
  Dropdown,
  IComboBoxOption,
  IDropdownOption,
  Label,
  TextField,
} from "@fluentui/react";
import * as moment from "moment";

export const Filters = (
  props: IFiltersProps
): React.ReactElement<unknown, React.JSXElementConstructor<unknown>> => {
  const statusNew = (
    event: React.MouseEvent<HTMLElement>,
    checked?: boolean
  ): void => {
    const arr: string[] = [...props.selectedStatus];
    if (checked) {
      arr.push("New");
    } else {
      arr.splice(props.selectedStatus.indexOf("New"), 1);
    }
    props.setSelectedStatus(arr);
  };
  const statusInProgress = (
    event: React.MouseEvent<HTMLElement>,
    checked?: boolean
  ): void => {
    const arr: string[] = [...props.selectedStatus];
    if (checked) {
      arr.push("In Progress");
    } else {
      arr.splice(props.selectedStatus.indexOf("In Progress"), 1);
    }
    props.setSelectedStatus(arr);
  };
  const statusApproved = (
    event: React.MouseEvent<HTMLElement>,
    checked?: boolean
  ): void => {
    const arr: string[] = [...props.selectedStatus];
    if (checked) {
      arr.push("Approved");
    } else {
      arr.splice(props.selectedStatus.indexOf("Approved"), 1);
    }
    props.setSelectedStatus(arr);
  };
  const statusRejected = (
    event: React.MouseEvent<HTMLElement>,
    checked?: boolean
  ): void => {
    const arr: string[] = [...props.selectedStatus];
    if (checked) {
      arr.push("Rejected");
    } else {
      arr.splice(props.selectedStatus.indexOf("Rejected"), 1);
    }
    props.setSelectedStatus(arr);
  };
  const onFormatDate = (date?: Date): string => {
    return !date
      ? ""
      : moment(
          `${date.getFullYear()}-${date.getMonth() + 1}-${date.getDate()}`,
          "YYYY-MM-DD"
        ).format("YYYY-MM-DD");
  };
  return (
    <>
      <div className={styles.filtersContainer}>
        <TextField
          label="Tags"
          placeholder="search"
          onChange={(e, value) => props.setTagsSearch(value)}
        />
        <TextField
          label="Title"
          placeholder="search"
          onChange={(e, value) => props.setTitleSearch(value)}
        />
      </div>
      <Label>Due Date</Label>
      <div className={styles.dateContainer}>
        <DatePicker
          placeholder="from"
          isMonthPickerVisible={false}
          onSelectDate={(date: Date) =>
            props.setDueDateStart(moment(date).format("YYYY-MM-DD"))
          }
          value={
            props.dueDateStart ? moment(props.dueDateStart).toDate() : undefined
          }
          formatDate={onFormatDate}
        />
        <DatePicker
          placeholder="to"
          isMonthPickerVisible={false}
          onSelectDate={(date: Date) =>
            props.setDueDateEnd(moment(date).format("YYYY-MM-DD"))
          }
          value={
            props.dueDateEnd ? moment(props.dueDateEnd).toDate() : undefined
          }
          formatDate={onFormatDate}
        />
      </div>
      <Label>Execution Date</Label>
      <div className={styles.dateContainer}>
        <DatePicker
          placeholder="from"
          isMonthPickerVisible={false}
          onSelectDate={(date: Date) =>
            props.setExecutionDateStart(moment(date).format("YYYY-MM-DD"))
          }
          value={
            props.executionDateStart
              ? moment(props.executionDateStart).toDate()
              : undefined
          }
          formatDate={onFormatDate}
        />
        <DatePicker
          placeholder="to"
          isMonthPickerVisible={false}
          onSelectDate={(date: Date) =>
            props.setExecutionDateEnd(moment(date).format("YYYY-MM-DD"))
          }
          value={
            props.executionDateEnd
              ? moment(props.executionDateEnd).toDate()
              : undefined
          }
          formatDate={onFormatDate}
        />
      </div>
      <div>
        {props.requestManagers !== undefined ? (
          <ComboBox
            className={styles.formField}
            label="Assigned manager"
            placeholder="select Assigned manager"
            options={props.requestManagers}
            autoComplete="on"
            onItemClick={(e, option: IComboBoxOption) =>
              typeof option.key !== "string"
                ? props.setManagerId(option.key)
                : null
            }
          />
        ) : null}
      </div>
      {props.requestTypes !== undefined ?
      <Dropdown
        className={styles.formField}
        label="Request Type"
        placeholder="select request type"
        onChange={(e, item: IDropdownOption) =>
          typeof item.key !== "string"
            ? props.setSelectedRequestTypeId(item.key)
            : null
        }
        options={props.requestTypes}
      /> : null}
      {props.requestAreaChoices !== undefined ? (
        <Dropdown
          className={styles.formField}
          label="Request area"
          placeholder="select request area"
          onChange={(e, item: IDropdownOption) =>
            props.setSelectedRequestAreaChoice(item.text)
          }
          options={props.requestAreaChoices}
        />
      ) : null}
      <div className={styles.toggleContainer}>
        <div className={styles.toggle}>
          <Toggle label="New" defaultChecked onChange={statusNew} />
        </div>
        <div className={styles.toggle}>
          <Toggle
            label="In Progress"
            defaultChecked
            onChange={statusInProgress}
          />
        </div>
        <div className={styles.toggle}>
          <Toggle label="Rejected" defaultChecked onChange={statusRejected} />
        </div>
        <div className={styles.toggle}>
          <Toggle label="Approved" defaultChecked onChange={statusApproved} />
        </div>
      </div>
    </>
  );
};
