import * as React from "react";
import styles from "./ListWebPart.module.scss";
import { IFiltersProps } from "./ListItems";
import {
  ComboBox,
  DatePicker,
  DefaultButton,
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
  const onFormatDate = (date?: Date): string => {
    return !date
      ? ""
      : moment(
          `${date.getFullYear()}-${date.getMonth() + 1}-${date.getDate()}`,
          "YYYY-MM-DD"
        ).format("YYYY-MM-DD");
  };
  function clearFilters(): void {
    props.setTitleSearch('')
    props.setTagsSearch('')
    props.setManagerId(undefined)
    props.setSelectedRequestTypeId(undefined)
    props.setSelectedRequestAreaChoice('')
    props.setSelectedStatus('')
    props.setDueDateStart(undefined)
    props.setDueDateEnd(undefined)
    props.setExecutionDateStart(undefined)
    props.setExecutionDateEnd(undefined)
  }
  return (
    <>
      <div className={styles.filtersContainer}>
        <TextField
          className={styles.filterField}
          label="Title"
          placeholder="search"
          onChange={(e, value) => props.setTitleSearch(value)}
        />
        <TextField
          className={styles.filterField}
          label="Tags"
          placeholder="search"
          onChange={(e, value) => props.setTagsSearch(value)}
        />
        {props.requestManagers !== undefined ? (
          <ComboBox
            className={styles.filterField}
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
        {props.requestTypes !== undefined ? (
          <Dropdown
            className={styles.filterField}
            label="Request Type"
            placeholder="select request type"
            onChange={(e, item: IDropdownOption) =>
              typeof item.key !== "string"
                ? props.setSelectedRequestTypeId(item.key)
                : null
            }
            options={props.requestTypes}
          />
        ) : null}
        {props.requestAreaChoices !== undefined ? (
          <Dropdown
            className={styles.filterField}
            label="Request area"
            placeholder="select request area"
            onChange={(e, item: IDropdownOption) =>
              props.setSelectedRequestAreaChoice(item.text)
            }
            options={props.requestAreaChoices}
          />
        ) : null}
        {props.statusChoices !== undefined ? (
          <Dropdown
            className={styles.filterField}
            label="Status"
            placeholder="select status"
            onChange={(e, item: IDropdownOption) =>
              props.setSelectedStatus(item.text)
            }
            options={props.statusChoices}
          />
        ) : null}
      </div>
      <Label>Due Date</Label>
      <div className={styles.dateContainer}>
        <DatePicker
          className={styles.filterField}
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
          className={styles.filterField}
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
          className={styles.filterField}
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
          className={styles.filterField}
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
        <DefaultButton className={styles.clearButton} onClick={clearFilters} text="Clear Filters" />
      </div>
      
    </>
  );
};
