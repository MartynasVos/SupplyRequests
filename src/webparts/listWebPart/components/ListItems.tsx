import * as React from "react";
import { IListItemsProps } from "./List";
import { EditRegular } from "@fluentui/react-icons";
import {
  TableBody,
  TableCell,
  TableRow,
  Table,
  TableHeader,
  TableHeaderCell,
  Button,
  useTableFeatures,
  TableColumnDefinition,
  createTableColumn,
  useTableSort,
  TableColumnId,
  SortDirection,
} from "@fluentui/react-components";
import { useBoolean } from "@fluentui/react-hooks";
import { IRequest } from "./List";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { RequestForm } from "./RequestForm";
import * as moment from "moment";
import { IDropdownOption } from "@fluentui/react/lib/Dropdown";
import { DefaultButton, type IComboBoxOption } from "@fluentui/react";
import styles from "./ListWebPart.module.scss";
import { Filters } from "./Filters";

export interface IRequestFormProps {
  context: WebPartContext;
  requestTypes: IDropdownOption[];
  taxonomy: IDropdownOption[];
  requestAreaChoices: IDropdownOption[] | undefined;
  setItems: React.Dispatch<React.SetStateAction<IRequest[]>>;
  hidePopup: () => void;
  isPopupVisible: boolean;
  requestManagers: IComboBoxOption[] | undefined;
  isRequestManager: boolean;
  currentItem: IRequest | undefined;
  getItems: () => Promise<IRequest[]>;
}

export interface IFiltersProps {
  selectedStatus: string | undefined;
  setSelectedStatus: React.Dispatch<React.SetStateAction<string>>;
  statusChoices: IDropdownOption<unknown>[] | undefined;
  setTagsSearch: React.Dispatch<React.SetStateAction<string | undefined>>;
  setTitleSearch: React.Dispatch<React.SetStateAction<string | undefined>>;
  dueDateStart: string | undefined;
  setDueDateStart: React.Dispatch<React.SetStateAction<string | undefined>>;
  dueDateEnd: string | undefined;
  setDueDateEnd: React.Dispatch<React.SetStateAction<string | undefined>>;
  executionDateStart: string | undefined;
  setExecutionDateStart: React.Dispatch<
    React.SetStateAction<string | undefined>
  >;
  executionDateEnd: string | undefined;
  setExecutionDateEnd: React.Dispatch<React.SetStateAction<string | undefined>>;
  requestManagers: IComboBoxOption[] | undefined;
  setManagerId: React.Dispatch<React.SetStateAction<number | undefined>>;
  requestAreaChoices: IDropdownOption<unknown>[] | undefined;
  requestTypes: IDropdownOption<unknown>[] | undefined;
  selectedRequestTypeId: number | undefined;
  setSelectedRequestTypeId: React.Dispatch<
    React.SetStateAction<number | undefined>
  >;
  selectedRequestAreaChoice: string | undefined;
  setSelectedRequestAreaChoice: React.Dispatch<
    React.SetStateAction<string | undefined>
  >;
}

const columnHeaders = [
  { columnKey: "buttons", label: "" },
  { columnKey: "title", label: "Title" },
  { columnKey: "status", label: "Status" },
  { columnKey: "assignedManager", label: "Assigned manager" },
  { columnKey: "dueDate", label: "Due date" },
  { columnKey: "executionDate", label: "Execution date" },
  { columnKey: "requestType", label: "Request type" },
  { columnKey: "requestArea", label: "Request area" },
  { columnKey: "tags", label: "Tags" },
];
const columns: TableColumnDefinition<IRequest>[] = [
  createTableColumn<IRequest>({
    columnId: "title",
    compare: (a, b) => {
      return a.Title.localeCompare(b.Title);
    },
  }),
  createTableColumn<IRequest>({
    columnId: "dueDate",
    compare: (a, b) => {
      return moment(a.DueDate)
        .format("YYYY-MM-DD")
        .localeCompare(moment(b.DueDate).format("YYYY-MM-DD"));
    },
  }),
  createTableColumn<IRequest>({
    columnId: "executionDate",
    compare: (a, b) => {
      return moment(a.ExecutionDate)
        .format("YYYY-MM-DD")
        .localeCompare(moment(b.ExecutionDate).format("YYYY-MM-DD"));
    },
  }),
];

export const ListItems = (
  props: IListItemsProps
): React.ReactElement<unknown, React.JSXElementConstructor<unknown>> => {
  const [isPopupVisible, { setTrue: showPopup, setFalse: hidePopup }] =
    useBoolean(false);
  const [currentItem, setCurrentItem] = React.useState<IRequest>();
  const [tagsSearch, setTagsSearch] = React.useState<string>();
  const [titleSearch, setTitleSearch] = React.useState<string>();
  const [dueDateStart, setDueDateStart] = React.useState<string>();
  const [dueDateEnd, setDueDateEnd] = React.useState<string>();
  const [executionDateStart, setExecutionDateStart] = React.useState<string>();
  const [executionDateEnd, setExecutionDateEnd] = React.useState<string>();
  const [managerId, setManagerId] = React.useState<number>();
  const [selectedRequestTypeId, setSelectedRequestTypeId] =
    React.useState<number>();
  const [selectedRequestAreaChoice, setSelectedRequestAreaChoice] =
    React.useState<string>();
  const [selectedStatus, setSelectedStatus] = React.useState<string>();

  const items = props.items;
  const {
    getRows,
    sort: { getSortDirection, toggleColumnSort, sort },
  } = useTableFeatures(
    {
      columns,
      items,
    },
    [
      useTableSort({
        defaultSortState: { sortColumn: "title", sortDirection: "ascending" },
      }),
    ]
  );

  const headerSortProps = (
    columnId: TableColumnId
  ): {
    onClick: (e: React.MouseEvent) => void;
    sortDirection: SortDirection | undefined;
  } => ({
    onClick: (e: React.MouseEvent) => {
      toggleColumnSort(e, columnId);
    },
    sortDirection: getSortDirection(columnId),
  });

  const rows = sort(getRows());
  return (
    <>
      <Filters
        selectedStatus={selectedStatus}
        setSelectedStatus={setSelectedStatus}
        statusChoices={props.statusChoices}
        setTagsSearch={setTagsSearch}
        setTitleSearch={setTitleSearch}
        dueDateStart={dueDateStart}
        setDueDateStart={setDueDateStart}
        dueDateEnd={dueDateEnd}
        setDueDateEnd={setDueDateEnd}
        executionDateStart={executionDateStart}
        setExecutionDateStart={setExecutionDateStart}
        executionDateEnd={executionDateEnd}
        setExecutionDateEnd={setExecutionDateEnd}
        requestManagers={props.requestManagers}
        setManagerId={setManagerId}
        requestAreaChoices={props.requestAreaChoices}
        requestTypes={props.requestTypes}
        selectedRequestTypeId={selectedRequestTypeId}
        setSelectedRequestTypeId={setSelectedRequestTypeId}
        selectedRequestAreaChoice={selectedRequestAreaChoice}
        setSelectedRequestAreaChoice={setSelectedRequestAreaChoice}
      />
      {!props.isRequestManager ? (
        <DefaultButton
          onClick={() => {
            showPopup();
            setCurrentItem(undefined);
          }}
          text="Create New Request"
        />
      ) : null}
      <Table arial-label="Default table" noNativeElements={true}>
        <TableHeader>
          <TableRow>
            {columnHeaders.map((column) => (
              <TableHeaderCell
                className={
                  column.columnKey === "buttons"
                    ? styles.tableButtons
                    : column.columnKey === "status" ? styles.statusCell
                    : column.columnKey === "dueDate" ? styles.dueDateCell
                    : column.columnKey === "executionDate" ? styles.executionDateCell
                    : column.columnKey === "requestArea" ? styles.requestAreaCell 
                    : column.columnKey === "requestType" ? styles.requestTypeCell 
                    : column.columnKey === "tags" ? styles.tagsCell
                    : undefined
                }
                {...headerSortProps(column.columnKey)}
                key={column.columnKey}
              >
                {column.label}
              </TableHeaderCell>
            ))}
          </TableRow>
        </TableHeader>
        <TableBody>
          {rows.map(({ item }) =>
            (!selectedStatus || selectedStatus === item.Status) &&
            (!selectedRequestAreaChoice ||
              selectedRequestAreaChoice === item.RequestArea) &&
            (!selectedRequestTypeId ||
              selectedRequestTypeId === item.RequestTypeId) &&
            (!managerId || managerId === item.Assigned_x0020_ManagerId) &&
            (!executionDateEnd ||
              moment(item.ExecutionDate).format("YYYY-MM-DD") <=
                executionDateEnd) &&
            (!executionDateStart ||
              moment(item.ExecutionDate).format("YYYY-MM-DD") >=
                executionDateStart) &&
            (!dueDateEnd ||
              moment(item.DueDate).format("YYYY-MM-DD") <= dueDateEnd) &&
            (!dueDateStart ||
              moment(item.DueDate).format("YYYY-MM-DD") >= dueDateStart) &&
            (!titleSearch ||
              item.Title.toLowerCase().indexOf(titleSearch.toLowerCase()) !==
                -1) &&
            (!tagsSearch ||
              item.Tags.filter(
                (tag: { Label: string }) =>
                  tag.Label.toLowerCase().indexOf(tagsSearch.toLowerCase()) !==
                  -1
              ).length !== 0) ? (
              <TableRow key={item.Id}>
                <TableCell className={styles.tableButtons}>
                  {item.Status === "New" ? (
                    <Button
                      onClick={() => {
                        showPopup();
                        setCurrentItem(item);
                      }}
                      icon={<EditRegular />}
                    />
                  ) : null}
                </TableCell>
                <TableCell>{item.Title}</TableCell>
                <TableCell className={styles.statusCell}>
                  {
                    <div
                      className={`${
                        item.Status === "New"
                          ? styles.new
                          : item.Status === "In Progress"
                          ? styles.inProgress
                          : item.Status === "Rejected"
                          ? styles.rejected
                          : styles.approved
                      } ${styles.status}`}
                    >
                      {item.Status}
                    </div>
                  }
                </TableCell>
                <TableCell>
                  {item.Assigned_x0020_ManagerId !== null &&
                  props.users[0] !== undefined
                    ? props.users.filter((user) => {
                        return user.Id === item.Assigned_x0020_ManagerId;
                      })[0].Title
                    : null}
                </TableCell>
                <TableCell className={styles.dueDateCell}>
                  {moment(item.DueDate).format("YYYY-MM-DD")}
                </TableCell>
                <TableCell className={styles.executionDateCell}>
                  {item.ExecutionDate !== null
                    ? moment(item.ExecutionDate).format("YYYY-MM-DD")
                    : "-"}
                </TableCell>
                
                <TableCell className={styles.requestTypeCell}>
                  {props.requestTypes[0] !== undefined
                    ? props.requestTypes.filter(
                        (type) => type.key === item.RequestTypeId
                      )[0].text
                    : null}
                </TableCell>
                <TableCell className={styles.requestAreaCell}>{item.RequestArea}</TableCell>
                <TableCell className={styles.tagsCell}>
                  {item.Tags.map(
                    (
                      tag: { Label: string },
                      index: React.Key | null | undefined
                    ) => {
                      return (
                        <div className={styles.tag} key={index}>
                          {tag.Label}
                        </div>
                      );
                    }
                  )}
                </TableCell>
              </TableRow>
            ) : null
          )}
        </TableBody>
      </Table>
      <RequestForm
        context={props.context}
        requestTypes={props.requestTypes}
        taxonomy={props.taxonomy}
        requestAreaChoices={props.requestAreaChoices}
        setItems={props.setItems}
        hidePopup={hidePopup}
        isPopupVisible={isPopupVisible}
        requestManagers={props.requestManagers}
        isRequestManager={props.isRequestManager}
        currentItem={currentItem}
        getItems={props.getItems}
      />
    </>
  );
};
