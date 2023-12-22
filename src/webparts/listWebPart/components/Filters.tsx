import * as React from 'react'
import { Toggle } from "@fluentui/react/lib/Toggle";
import styles from "./ListWebPart.module.scss";
import { IFiltersProps } from './ListItems';
import { TextField } from '@fluentui/react';

export const Filters = (props: IFiltersProps): React.ReactElement<unknown, React.JSXElementConstructor<unknown>> => {
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
  return (
    <div className={styles.toggleBox}>
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
    <TextField
        label="Tags"
        placeholder="search"
        onChange={(e, value) => props.setTagsSearch(value)}
      />
  </div>
  )
}
