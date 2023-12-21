import * as React from "react";
import { IListItemsProps } from "./List";
import { EditRegular, DeleteRegular } from "@fluentui/react-icons";
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
import { EditItems } from "./EditItem";
import * as moment from "moment";
import { SPFx, spfi } from "@pnp/sp";
import { IDropdownOption } from "@fluentui/react/lib/Dropdown";
import type { IComboBoxOption } from "@fluentui/react";

const columnHeaders = [
  { columnKey: "title", label: "Title" },
  { columnKey: "dueDate", label: "Due date" },
  { columnKey: "executionDate", label: "Execution date" },
  { columnKey: "requestType", label: "Request type" },
  { columnKey: "requestArea", label: "Request area" },
  { columnKey: "assignedManager", label: "Assigned manager" },
  { columnKey: "tags", label: "Tags" },
  { columnKey: "status", label: "Status" },
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
      return moment(a.ExecutionDate).format("YYYY-MM-DD").localeCompare(moment(b.ExecutionDate).format("YYYY-MM-DD"));
    },
  }),

  createTableColumn<IRequest>({
    columnId: "requestArea",
    compare: (a, b) => {
      return a.RequestArea.localeCompare(b.RequestArea);
    },
  }),
];

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
  currentItem: IRequest | undefined;
}

export const ListItems = (
  props: IListItemsProps
): React.ReactElement<unknown, React.JSXElementConstructor<unknown>> => {
  const [isPopupVisible, { setTrue: showPopup, setFalse: hidePopup }] =
    useBoolean(false);

  const [currentItem, setCurrentItem] = React.useState<IRequest>();

  function edit(item: IRequest): void {
    if (item.Status !== "New") {
      return alert("Can only edit request in status new");
    }
    showPopup();
    setCurrentItem(item);
  }
  function deleteItemFunction(item: IRequest): void {
    if (item.Status !== "New") {
      return alert("Can only delete request in status new");
    }
    if (!confirm("Are you sure you want to delete this request?")) {
      return;
    }
    const deleteItem = async (): Promise<void> => {
      const sp = spfi().using(SPFx(props.context));
      const list = sp.web.lists.getByTitle("Requests");
      const i = await list.items.getById(item.Id).delete();
      console.log(i);
    };
    deleteItem().then(
      () => {
        const getItems = async (): Promise<IRequest[]> => {
          const sp = spfi().using(SPFx(props.context));
          const items = await sp.web.lists.getByTitle("Requests").items();
          return items;
        };
        getItems().then(
          (result) => {
            props.setItems(result);
          },
          () => {
            return;
          }
        );
      },
      () => {
        return;
      }
    );
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
          {rows.map(({ item }) => (
            <TableRow key={item.Title}>
              <TableCell>{item.Title}</TableCell>
              <TableCell>{moment(item.DueDate).format("YYYY-MM-DD")}</TableCell>
              <TableCell>
                <TableCellLayout truncate={true}>
                  {item.ExecutionDate !== null
                    ? moment(item.ExecutionDate).format("YYYY-MM-DD")
                    : "-"}
                </TableCellLayout>
              </TableCell>
              <TableCell>
                {props.requestTypes.map((type) => {
                  if (type.key === item.RequestTypeId) {
                    return type.text;
                  }
                })}
              </TableCell>
              <TableCell>{item.RequestArea}</TableCell>
              <TableCell>
                {props.users.map((user) => {
                  if (user.Id === item.Assigned_x0020_ManagerId) {
                    return user.Title;
                  }
                })}
              </TableCell>
              <TableCell>
                {item.Tags.map(
                  (
                    tag: { Label: string },
                    index: React.Key | null | undefined
                  ) => {
                    return <div key={index}>{tag.Label}</div>;
                  }
                )}
              </TableCell>
              <TableCell>{item.Status}</TableCell>
              <TableCell>
                <Button
                  onClick={() => {
                    edit(item);
                  }}
                  icon={<EditRegular />}
                />
                <Button
                  onClick={() => deleteItemFunction(item)}
                  icon={<DeleteRegular />}
                />
              </TableCell>
            </TableRow>
          ))}
        </TableBody>
      </Table>
      <EditItems
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
      />
    </>
  );
};
