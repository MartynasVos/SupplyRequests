import * as React from "react";
import { IEditItemProps } from "./ListItems";
import {
  DatePicker,
  DefaultButton,
  FocusTrapZone,
  Layer,
  Overlay,
  Popup,
  TextField,
  addDays,
  mergeStyleSets,
  ComboBox,
  IComboBoxOption,
  IComboBox
} from "@fluentui/react";
import {
  Dropdown,
  IDropdownOption,
} from "@fluentui/react/lib/Dropdown";

import { SPFx, spfi } from "@pnp/sp";
import { IRequest } from "./List";
import * as moment from "moment";

const styles = mergeStyleSets({
  root: {
    background: "rgba(0, 0, 0, 0.2)",
    bottom: "0",
    left: "0",
    position: "fixed",
    right: "0",
    top: "0",
  },
  content: {
    background: "white",
    left: "50%",
    maxWidth: "640px",
    width: "100%",
    padding: "2em",
    position: "absolute",
    top: "50%",
    transform: "translate(-50%, -50%)",
    display: "grid",
    justifyContent: "center",
  },
  formField: {
    width: 300,
    marginBottom: "40px",
  },
});

export const EditItems = (
  props: IEditItemProps
): React.ReactElement<unknown, React.JSXElementConstructor<unknown>> => {
  const [selectedManagerId, setSelectedManagerId] = React.useState<number>();
  const [selectedDate, setSelectedDate] = React.useState<Date>();
  const [selectedRequestTypeId, setSelectedRequestTypeId] =
    React.useState<number>();
  const [selectedRequestAreaChoice, setSelectedRequestAreaChoice] =
    React.useState<string>();
  const [selectedTagsIds, setSelectedTagsIds] = React.useState<string[]>([]);
  const [currentItemId, setCurrentItemId] = React.useState<number>(0);

  React.useEffect(() => {
    setData();
  }, [props.currentItem]);

  function setData(): void {
    setSelectedDate(moment(props.currentItem?.DueDate).toDate());
    setSelectedRequestTypeId(props.currentItem?.RequestTypeId);
    setSelectedRequestAreaChoice(props.currentItem?.RequestArea);
    const tagIds: string[] = [];
    props.currentItem?.Tags.map((tag: { TermGuid: string }) => {
      tagIds.push(tag.TermGuid);
    });
    setSelectedTagsIds(tagIds);
    if (props.currentItem !== undefined) {
      setCurrentItemId(props.currentItem?.Id);
    }
    setSelectedManagerId(undefined)
  }

  function editItemFunction(): void {
    let title = props.currentItem?.Title;
    if (document.getElementById("title") !== null) {
      title = (document.getElementById("title") as HTMLInputElement).value;
    }
    let description = props.currentItem?.Description;
    if (document.getElementById("description") !== null) {
      description = (document.getElementById("description") as HTMLInputElement)
        .value;
    }
    let status = "New"
      if (props.isRequestManager) {
        status = 'In Progress'
        if (selectedManagerId === undefined) {
         return alert('Assigned Manager field is mandatory')
        }
      }
    const editItem = async (): Promise<void> => {
      const sp = spfi().using(SPFx(props.context));
      const list = sp.web.lists.getByTitle("Requests");
      const i = await list.items.getById(currentItemId).update({
        Title: title,
        Description: description,
        DueDate: selectedDate,
        Assigned_x0020_ManagerId: selectedManagerId,
        RequestTypeId: selectedRequestTypeId,
        RequestArea: selectedRequestAreaChoice,
        Status: status,
      });
      const fields = await sp.web.lists
        .getByTitle("Requests")
        .fields.filter("Title eq 'Tags_0'")
        .select("Title", "InternalName")();
      const updateTags: { [key: string]: unknown } = {};
      updateTags[fields[0].InternalName] = selectedTagsIds.join(";");
      await i.item.update(updateTags);
      console.log(i);
    };

    editItem().then(
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
    props.hidePopup();
    
  }
  const setRequestType = (
    event: React.FormEvent<HTMLDivElement>,
    item: IDropdownOption
  ): void => {
    if (typeof item.key !== "string") {
      setSelectedRequestTypeId(item.key);
    }
  };
  const setTags = (
    event: React.FormEvent<HTMLDivElement>,
    item: IDropdownOption
  ): void => {
    if (typeof item.key === "string")
      if (selectedTagsIds.indexOf(item.key) === -1) {
        selectedTagsIds.push(item.key);
      } else {
        selectedTagsIds.splice(selectedTagsIds.indexOf(item.key), 1);
      }
  };

  const onFormatDate = (date?: Date): string => {
    return !date
      ? ""
      : moment(
          `${date.getFullYear()}-${date.getMonth() + 1}-${date.getDate()}`,
          "YYYY-MM-DD"
        ).format("YYYY-MM-DD");
  };

  const setManager = (
    event: React.FormEvent<IComboBox>, option?: IComboBoxOption | undefined, index?: number | undefined
  ): void => {
    if (option !== undefined && typeof option.key === 'number') {
      setSelectedManagerId(option.key)
    }
  };
  return (
    <>
      {props.isPopupVisible && (
        <Layer>
          <Popup
            className={styles.root}
            role="dialog"
            aria-modal="true"
            onDismiss={props.hidePopup}
            enableAriaHiddenSiblings={true}
          >
            <Overlay onClick={props.hidePopup} />
            <FocusTrapZone>
              <div role="document" className={styles.content}>
                    <TextField
                      label="Title"
                      id="title"
                      required
                      disabled={props.isRequestManager ? true : false}
                      defaultValue={props.currentItem?.Title}
                    />
                    <TextField
                      label="Description"
                      id="description"
                      required
                      multiline
                      rows={5}
                      disabled={props.isRequestManager ? true : false}
                      defaultValue={props.currentItem?.Description}
                    />
                
                {props.isRequestManager ? (
                  <div>
                    {props.requestManagers !== undefined ? (
                      <ComboBox 
                      className={styles.formField}
                      label="Assign a manager"
                      required
                      options={props.requestManagers}
                      autoComplete="on"
                      onItemClick={setManager} />
                    ) : null}
                  </div>
                ) : null}
                <div>
                  <Dropdown
                    className={styles.formField}
                    label="Request Type"
                    defaultSelectedKey={selectedRequestTypeId}
                    onChange={setRequestType}
                    options={props.requestTypes}
                  />
                </div>
                <div>
                  {props.requestAreaChoices !== undefined ? (
                    <Dropdown
                      className={styles.formField}
                      label="Request area"
                      defaultSelectedKey={selectedRequestAreaChoice}
                      onChange={(e, item: IDropdownOption) => setSelectedRequestAreaChoice(item.text)}
                      options={props.requestAreaChoices}
                    />
                  ) : null}
                </div>
                  <div>
                    <DatePicker
                      id="dueDate"
                      className={styles.formField}
                      label="Due Date"
                      isMonthPickerVisible={false}
                      minDate={addDays(new Date(), 3)}
                      onSelectDate={(date: Date) =>
                        setSelectedDate(moment(date, "YYYY-MM-DD").toDate())
                      }
                      value={selectedDate}
                      formatDate={onFormatDate}
                      disabled={props.isRequestManager ? true : false}
                    />
                  </div>
               
                <div>
                  <Dropdown
                    className={styles.formField}
                    label="Tags"
                    defaultSelectedKeys={selectedTagsIds}
                    onChange={setTags}
                    options={props.taxonomy}
                    multiSelect
                  />
                </div>
                <div>
                  <DefaultButton
                    onClick={() => {
                      editItemFunction();
                    }}
                  >
                    {props.isRequestManager
                      ? "Send to delivery department"
                      : "Edit"}
                  </DefaultButton>
                  <DefaultButton
                    onClick={() => {
                      props.hidePopup();
                      setData();
                    }}
                  >
                    Cancel
                  </DefaultButton>
                </div>
              </div>
            </FocusTrapZone>
          </Popup>
        </Layer>
      )}
    </>
  );
};
