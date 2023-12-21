import * as React from "react";
import {
  mergeStyleSets,
  DefaultButton,
  FocusTrapZone,
  Layer,
  Overlay,
  Popup,
  DatePicker,
  addDays,
  TextField,
} from "@fluentui/react";

import {
  Dropdown,
  IDropdownOption,
} from "@fluentui/react/lib/Dropdown";

import { useBoolean } from "@fluentui/react-hooks";
import { IFormProps, IRequest } from "./List";
import { SPFx, spfi } from "@pnp/sp";
import "@pnp/sp/fields";

import { IItemAddResult } from "@pnp/sp/items";
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
    width: "300px",
    marginBottom: "40px",
  },
  primaryBtn: {
    ":hover": {
      backgroundColor: "blue",
    },
    ":active": {
      backgroundColor: "blue",
    },
    backgroundColor: "blue",
    color: "white",
  },
});

export const FormModalBox = (
  props: IFormProps
): React.ReactElement<unknown, React.JSXElementConstructor<unknown>> => {
  const [isPopupVisible, { setTrue: showPopup, setFalse: hidePopup }] =
    useBoolean(false);

  const [selectedTagsIds] = React.useState<string[]>([]);
  const [selectedDate, setSelectedDate] = React.useState<Date>();
  const [selectedRequestAreaChoice, setSelectedRequestAreaChoice] =
    React.useState<string>();
  const [selectedRequestTypeId, setSelectedRequestTypeId] =
    React.useState<number>();

  function addItemFunction(): void {
    const title = (document.getElementById("title") as HTMLInputElement).value;
    const description = (
      document.getElementById("description") as HTMLInputElement
    ).value;
    if (title === "") {
      return alert("Title field is mandatory");
    }
    if (description === "") {
      return alert("Description field is mandatory");
    }
    if (selectedDate === undefined) {
      return alert("Due date field is mandatory");
    }
    if (selectedRequestTypeId === undefined) {
      return alert("Request type field is mandatory");
    }
    const sp = spfi().using(SPFx(props.context));
    const addItem = async (): Promise<void> => {
      const iar: IItemAddResult = await sp.web.lists
        .getByTitle("Requests")
        .items.add({
          Title: title,
          Description: description,
          DueDate: selectedDate,
          RequestTypeId: selectedRequestTypeId,
          RequestArea: selectedRequestAreaChoice,
        });
      const fields = await sp.web.lists
        .getByTitle("Requests")
        .fields.filter("Title eq 'Tags_0'")
        .select("Title", "InternalName")();
      const updateTags: { [key: string]: unknown } = {};
      updateTags[fields[0].InternalName] = selectedTagsIds.join(";");
      await iar.item.update(updateTags);
    };
    addItem().then(
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
    hidePopup();
    setSelectedDate(undefined);
    setSelectedRequestTypeId(undefined);
    setSelectedRequestAreaChoice(undefined);
  }
  const onFormatDate = (date?: Date): string => {
    return !date
      ? ""
      : moment(
          `${date.getFullYear()}-${date.getMonth() + 1}-${date.getDate()}`,
          "YYYY-MM-DD"
        ).format("YYYY-MM-DD");
  };
  const setRequestType = (
    event: React.FormEvent<HTMLDivElement>,
    item: IDropdownOption
  ): void => {
    if (typeof item.key !== "string") {
      setSelectedRequestTypeId(item.key);
    }
  };
  const setRequestAreaChoice = (
    event: React.FormEvent<HTMLDivElement>,
    item: IDropdownOption
  ): void => {
    setSelectedRequestAreaChoice(item.text);
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
  return (
    <>
      <DefaultButton
        onClick={() => {
          if (props.isRequestManager) {
            return alert("Managers can't create new requests");
          } else {
            showPopup();
          }
        }}
        text="Add Request"
      />
      {isPopupVisible && (
        <Layer>
          <Popup
            className={styles.root}
            role="dialog"
            aria-modal="true"
            onDismiss={hidePopup}
            enableAriaHiddenSiblings={true}
          >
            <Overlay onClick={hidePopup} />
            <FocusTrapZone>
              <div role="document" className={styles.content}>
                <TextField label="Title" id="title" required />
                <TextField
                  label="Description"
                  id="description"
                  required
                  multiline
                  rows={5}
                />
                <div>
                  <Dropdown
                    className={styles.formField}
                    label="Request Type"
                    onChange={setRequestType}
                    options={props.requestTypes}
                  />
                </div>
                <div>
                  {props.requestAreaChoices !== undefined ? (
                    <Dropdown
                      className={styles.formField}
                      label="Request area"
                      onChange={setRequestAreaChoice}
                      options={props.requestAreaChoices}
                    />
                  ) : null}
                </div>
                <div>
                  <DatePicker
                    id="dueDate"
                    placeholder="Select a date..."
                    className={styles.formField}
                    label="Due Date"
                    isMonthPickerVisible={false}
                    minDate={addDays(new Date(), 3)}
                    onSelectDate={(date: Date) => setSelectedDate(date)}
                    value={selectedDate}
                    formatDate={onFormatDate}
                  />
                </div>
                <div>
                  <Dropdown
                    className={styles.formField}
                    label="Tags"
                    onChange={setTags}
                    options={props.taxonomy}
                    multiSelect
                  />
                </div>
                <div>
                  <DefaultButton
                    className={styles.primaryBtn}
                    onClick={() => {
                      addItemFunction();
                    }}
                  >
                    Save
                  </DefaultButton>
                  <DefaultButton
                    onClick={() => {
                      hidePopup();
                      setSelectedDate(undefined);
                      setSelectedRequestTypeId(undefined);
                      setSelectedRequestAreaChoice(undefined);
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
