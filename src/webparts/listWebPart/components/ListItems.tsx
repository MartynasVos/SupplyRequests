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
  TableCellLayout,
} from "@fluentui/react-components";
import { useBoolean } from "@fluentui/react-hooks";
import { IRequest } from "./List";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { EditItem } from "./EditItem";
import * as moment from "moment";
import { IDropdownOption } from "@fluentui/react/lib/Dropdown";
import { type IComboBoxOption } from "@fluentui/react";
import styles from "./ListWebPart.module.scss";
import { Filters } from "./Filters";

export interface IEditItemProps {
  context: WebPartContext;
  requestTypes: IDropdownOption[];
  taxonomy: IDropdownOption[];
  requestAreaChoices: IDropdownOption[] | undefined;
  setItems: React.Dispatch<React.SetStateAction<IRequest[]>>;
  hidePopup: () => void;
  isPopupVisible: boolean;
  requestManagers: IComboBoxOption[] | undefined;
  isRequestManager: boolean;
  currentItem: IRequest;
  getItems: () => Promise<IRequest[]>;
}

export interface IFiltersProps {
  selectedStatus: string[];
  setSelectedStatus: React.Dispatch<React.SetStateAction<string[]>>;
  setTagsSearch: React.Dispatch<React.SetStateAction<string | undefined>>;
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

  const [currentItem, setCurrentItem] = React.useState<IRequest>(
    props.items[0]
  );
  const [tagsSearch, setTagsSearch] = React.useState<string>();
  const [selectedStatus, setSelectedStatus] = React.useState([
    "New",
    "In Progress",
    "Rejected",
    "Approved",
  ]);

  function edit(item: IRequest): void {
    showPopup();
    setCurrentItem(item);
  }

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
    setTagsSearch={setTagsSearch}
    />
      <Table arial-label="Default table">
        <TableHeader>
          <TableRow>
            {columnHeaders.map((column) => (
              <TableHeaderCell
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
            (!tagsSearch ||
              item.Tags.filter(
                (tag: { Label: string }) =>
                  tag.Label.toLowerCase().indexOf(tagsSearch.toLowerCase()) !==
                  -1
              ).length !== 0) &&
            selectedStatus.indexOf(item.Status) !== -1 ? (
              <TableRow key={item.Id}>
                <TableCell style={{maxWidth: '20px'}}>
                  {item.Status === "New" ? (
                      <Button
                        onClick={() => {
                          edit(item);
                        }}
                        icon={<EditRegular />}
                      />
                  ) : null}
                </TableCell>
                <TableCell>{item.Title}</TableCell>
                <TableCell>
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
                  {item.Assigned_x0020_ManagerId !== null
                    ? props.users.filter((user) => {
                        return user.Id === item.Assigned_x0020_ManagerId;
                      })[0].Title
                    : null}
                </TableCell>
                <TableCell>
                  {moment(item.DueDate).format("YYYY-MM-DD")}
                </TableCell>
                <TableCell>
                  {item.ExecutionDate !== null
                    ? moment(item.ExecutionDate).format("YYYY-MM-DD")
                    : "-"}
                </TableCell>
                <TableCell>
                  {
                    props.requestTypes.filter(
                      (type) => type.key === item.RequestTypeId
                    )[0].text
                  }
                </TableCell>
                <TableCell>{item.RequestArea}</TableCell>
                <TableCell>
                  <TableCellLayout>
                    <div className={styles.tagsCell}>
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
                    </div>
                  </TableCellLayout>
                </TableCell>
              </TableRow>
            ) : null
          )}
        </TableBody>
      </Table>
      <EditItem
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
