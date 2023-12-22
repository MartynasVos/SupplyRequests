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
  ComboBox,
  IComboBoxOption,
  IComboBox,
} from "@fluentui/react";
import { Dropdown, IDropdownOption } from "@fluentui/react/lib/Dropdown";
import { useBoolean } from "@fluentui/react-hooks";
import { IFormProps } from "./List";
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

  const [selectedTagsIds, setSelectedTagsIds] = React.useState<string[]>([]);
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
        props.getItems().then(
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
    setSelectedTagsIds([])
  }

  const onFormatDate = (date?: Date): string => {
    return !date
      ? ""
      : moment(
          `${date.getFullYear()}-${date.getMonth() + 1}-${date.getDate()}`,
          "YYYY-MM-DD"
        ).format("YYYY-MM-DD");
  };

  const setTags = (event: React.FormEvent<IComboBox>, option?: IComboBoxOption): void => {
    if (option && typeof option.key === "string")
      if (selectedTagsIds.indexOf(option.key) === -1) {
        selectedTagsIds.push(option.key);
      } else {
        selectedTagsIds.splice(selectedTagsIds.indexOf(option.key), 1);
      }
  };

  return (
    <>
      <DefaultButton
        onClick={() => {
            showPopup();
        }}
        text="Create New Request"
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
                  className={styles.formField}
                  label="Due Date"
                  isRequired
                  isMonthPickerVisible={false}
                  minDate={addDays(new Date(), 3)}
                  onSelectDate={(date: Date) => setSelectedDate(date)}
                  value={selectedDate}
                  formatDate={onFormatDate}
                />
                <Dropdown
                  className={styles.formField}
                  label="Request Type"
                  required
                  onChange={(e, item: IDropdownOption) =>
                    typeof item.key !== "string"
                      ? setSelectedRequestTypeId(item.key)
                      : null
                  }
                  options={props.requestTypes}
                />
                {props.requestAreaChoices !== undefined ? (
                  <Dropdown
                    className={styles.formField}
                    label="Request area"
                    onChange={(e, item: IDropdownOption) =>
                      setSelectedRequestAreaChoice(item.text)
                    }
                    options={props.requestAreaChoices}
                  />
                ) : null}
                <ComboBox
                  className={styles.formField}
                  multiSelect
                  label="Tags"
                  options={props.taxonomy}
                  autoComplete="on"
                  onChange={setTags}
                />
                <div className={styles.formButtonsContainer}>
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
                      setSelectedTagsIds([]);
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
