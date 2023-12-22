import * as React from "react";
import {
  DefaultButton,
  FocusTrapZone,
  Layer,
  Overlay,
  Popup,
  DatePicker,
  addDays,
  TextField,
} from "@fluentui/react";
import { Dropdown, IDropdownOption } from "@fluentui/react/lib/Dropdown";
import { useBoolean } from "@fluentui/react-hooks";
import { IFormProps, IRequest } from "./List";
import { SPFx, spfi } from "@pnp/sp";
import "@pnp/sp/fields";
import { IItemAddResult } from "@pnp/sp/items";
import * as moment from "moment";
import styles from "./ListWebPart.module.scss";

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
            className={styles.modalBox}
            role="dialog"
            aria-modal="true"
            onDismiss={hidePopup}
            enableAriaHiddenSiblings={true}
          >
            <Overlay onClick={hidePopup} />
            <FocusTrapZone>
              <div role="document" className={styles.modalContent}>
                <TextField label="Title" id="title" required />
                <TextField
                  label="Description"
                  id="description"
                  required
                  multiline
                  rows={5}
                />
                <DatePicker
                  placeholder="Select a date..."
                  className={styles.modalFormField}
                  label="Due Date"
                  isRequired
                  isMonthPickerVisible={false}
                  minDate={addDays(new Date(), 3)}
                  onSelectDate={(date: Date) => setSelectedDate(date)}
                  value={selectedDate}
                  formatDate={onFormatDate}
                />
                <Dropdown
                  className={styles.modalFormField}
                  label="Request Type"
                  required
                  onChange={(e, item: IDropdownOption) =>
                    typeof item.key !== "string"
                      ? setSelectedRequestTypeId(item.key)
                      : null}
                  options={props.requestTypes}
                />
                {props.requestAreaChoices !== undefined ? (
                  <Dropdown
                    className={styles.modalFormField}
                    label="Request area"
                    onChange={(e, item: IDropdownOption) =>
                      setSelectedRequestAreaChoice(item.text)}
                    options={props.requestAreaChoices}
                  />
                ) : null}
                <Dropdown
                  className={styles.modalFormField}
                  label="Tags"
                  onChange={setTags}
                  options={props.taxonomy}
                  multiSelect
                />
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
