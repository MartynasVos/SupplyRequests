import * as React from "react";
import { IRequestFormProps } from "./ListItems";
import {
  DatePicker,
  DefaultButton,
  FocusTrapZone,
  Layer,
  Overlay,
  Popup,
  TextField,
  addDays,
  ComboBox,
  IComboBoxOption,
  PrimaryButton,
  IComboBox,
  Dialog,
  DialogType,
  DialogFooter,
} from "@fluentui/react";
import { Dropdown, IDropdownOption } from "@fluentui/react/lib/Dropdown";
import { SPFx, spfi } from "@pnp/sp";
import { IRequest } from "./List";
import * as moment from "moment";
import styles from "./ListWebPart.module.scss";
import { DeleteItem } from "./DeleteItem";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IItemAddResult } from "@pnp/sp/items";
import { useBoolean } from "@fluentui/react-hooks";

export interface IDeleteItemProps {
  context: WebPartContext;
  setItems: React.Dispatch<React.SetStateAction<IRequest[]>>;
  hidePopup: () => void;
  currentItem: IRequest;
  getItems: () => Promise<IRequest[]>;
}

export const RequestForm = (
  props: IRequestFormProps
): React.ReactElement<unknown, React.JSXElementConstructor<unknown>> => {
  const [hideDialog, { toggle: toggleHideDialog }] = useBoolean(true);
  const [errorMessage, setErrorMessage] = React.useState<string>('');
  const dialogContentProps = {
    type: DialogType.normal,
    title: <div style={{textAlign: "center", color: "red", paddingLeft: '22px'}}>{errorMessage}</div>,
  };
  const [selectedManagerId, setSelectedManagerId] = React.useState<number>();
  const [selectedDate, setSelectedDate] = React.useState<Date>();
  const [selectedRequestTypeId, setSelectedRequestTypeId] =
    React.useState<number>();
  const [selectedRequestAreaChoice, setSelectedRequestAreaChoice] =
    React.useState<string>();
  const [selectedTagsIds, setSelectedTagsIds] = React.useState<string[]>([]);

  React.useEffect(() => {
    setData();
  }, [props.currentItem]);

  function setData(): void {
    setSelectedDate(
      props.currentItem ? moment(props.currentItem.DueDate).toDate() : undefined
    );
    setSelectedRequestTypeId(
      props.currentItem ? props.currentItem.RequestTypeId : undefined
    );
    setSelectedRequestAreaChoice(
      props.currentItem ? props.currentItem.RequestArea : undefined
    );
    const tagIds: string[] = [];
    props.currentItem?.Tags.map((tag: { TermGuid: string }) => {
      tagIds.push(tag.TermGuid);
    });
    setSelectedTagsIds(tagIds);
    setSelectedManagerId(props.currentItem?.Assigned_x0020_ManagerId);
  }
  function addItemFunction(): void {
    const title = (document.getElementById("title") as HTMLInputElement).value;
    const description = (
      document.getElementById("description") as HTMLInputElement
    ).value;
    if (title === "") {
      setErrorMessage('Title Field is mandatory')
      return toggleHideDialog();
    }
    if (description === "") {
      setErrorMessage('Description Field is mandatory')
      return toggleHideDialog();
    }
    if (selectedDate === undefined) {
      setErrorMessage('Due date Field is mandatory')
      return toggleHideDialog();
    }
    if (selectedRequestTypeId === undefined) {
      setErrorMessage('Request type Field is mandatory')
      return toggleHideDialog();
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
    props.hidePopup();
    setSelectedDate(undefined);
    setSelectedRequestTypeId(undefined);
    setSelectedRequestAreaChoice(undefined);
    setSelectedManagerId(undefined);
    setSelectedTagsIds([]);
  }
  function editItemFunction(): void {
    const title = (document.getElementById("title") as HTMLInputElement).value;
    const description = (
      document.getElementById("description") as HTMLInputElement
    ).value;
    if (title === "") {
      setErrorMessage('Title Field is mandatory')
      return toggleHideDialog();
    }
    if (description === "") {
      setErrorMessage('Description Field is mandatory')
      return toggleHideDialog();
    }
    if (selectedDate === undefined) {
      setErrorMessage('Due date Field is mandatory')
      return toggleHideDialog();
    }
    if (selectedRequestTypeId === undefined) {
      setErrorMessage('Request type Field is mandatory')
      return toggleHideDialog();
    }
    if (props.isRequestManager) {
      if (selectedManagerId === null) {
        setErrorMessage('Assigned manager field is mandatory')
        return toggleHideDialog();
      }
    }
    const editItem = async (): Promise<void> => {
      const sp = spfi().using(SPFx(props.context));
      const list = sp.web.lists.getByTitle("Requests");
      if (props.currentItem !== undefined) {
        const i = await list.items.getById(props.currentItem.Id).update({
          Title: title,
          Description: description,
          DueDate: selectedDate,
          Assigned_x0020_ManagerId: selectedManagerId,
          RequestTypeId: selectedRequestTypeId,
          RequestArea: selectedRequestAreaChoice,
        });
        const fields = await sp.web.lists
          .getByTitle("Requests")
          .fields.filter("Title eq 'Tags_0'")
          .select("Title", "InternalName")();
        const updateTags: { [key: string]: unknown } = {};
        updateTags[fields[0].InternalName] = selectedTagsIds.join(";");
        await i.item.update(updateTags);
        console.log(i);
      }
    };
    editItem().then(
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
    props.hidePopup();
  }
  function sendToDeliveryDepartment(): void {
    if (props.isRequestManager) {
      if (selectedManagerId === null) {
        setErrorMessage('Assigned manager field is mandatory')
        return toggleHideDialog();
      }
    }
    const editItem = async (): Promise<void> => {
      const sp = spfi().using(SPFx(props.context));
      const list = sp.web.lists.getByTitle("Requests");
      if (props.currentItem !== undefined) {
        const i = await list.items.getById(props.currentItem.Id).update({
          Assigned_x0020_ManagerId: selectedManagerId,
          RequestTypeId: selectedRequestTypeId,
          RequestArea: selectedRequestAreaChoice,
          Status: "In Progress",
        });
        console.log(i);
      }
    };
    editItem().then(
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
    props.hidePopup();
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
    event: React.FormEvent<IComboBox>,
    option?: IComboBoxOption
  ): void => {
    if (option && typeof option.key === "string")
      if (selectedTagsIds.indexOf(option.key) === -1) {
        selectedTagsIds.push(option.key);
      } else {
        selectedTagsIds.splice(selectedTagsIds.indexOf(option.key), 1);
      }
  };
  return (
    <>
      {props.isPopupVisible && (
        <Layer>
          <Popup
            className={styles.modalBox}
            role="dialog"
            aria-modal="true"
            onDismiss={props.hidePopup}
            enableAriaHiddenSiblings={true}
          >
            <Overlay onClick={props.hidePopup} />
            <FocusTrapZone>
              <div role="document" className={styles.modalContent}>
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
                <DatePicker
                  className={styles.formField}
                  placeholder="Select a date..."
                  label="Due Date"
                  isRequired
                  isMonthPickerVisible={false}
                  minDate={addDays(new Date(), 3)}
                  onSelectDate={(date: Date) =>
                    setSelectedDate(moment(date, "YYYY-MM-DD").toDate())
                  }
                  value={selectedDate}
                  formatDate={onFormatDate}
                  disabled={props.isRequestManager ? true : false}
                />
                {props.isRequestManager ? (
                  props.requestManagers !== undefined ? (
                    <ComboBox
                      className={styles.formField}
                      label="Assigned manager"
                      placeholder="Assign a manger"
                      required
                      options={props.requestManagers}
                      autoComplete="on"
                      onItemClick={(e, option: IComboBoxOption) =>
                        typeof option.key !== "string"
                          ? setSelectedManagerId(option.key)
                          : null
                      }
                      defaultSelectedKey={selectedManagerId}
                    />
                  ) : null
                ) : null}
                <Dropdown
                  className={styles.formField}
                  label="Request Type"
                  placeholder="select request type"
                  required
                  defaultSelectedKey={selectedRequestTypeId}
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
                    placeholder="select request area"
                    defaultSelectedKey={selectedRequestAreaChoice}
                    onChange={(e, item: IDropdownOption) =>
                      setSelectedRequestAreaChoice(item.text)
                    }
                    options={props.requestAreaChoices}
                    selectedKey={
                      selectedRequestAreaChoice === undefined
                        ? null
                        : selectedRequestAreaChoice
                    }
                  />
                ) : null}
                <ComboBox
                  multiSelect
                  placeholder="type to add tags"
                  className={styles.formField}
                  label="Tags"
                  options={props.taxonomy}
                  autoComplete="on"
                  onChange={setTags}
                  defaultSelectedKey={selectedTagsIds}
                />
                <div className={styles.formButtonsContainer}>
                  {props.currentItem !== undefined ? (
                    <PrimaryButton
                      className={styles.primaryBtn}
                      onClick={() => {
                        editItemFunction();
                      }}
                    >
                      Update
                    </PrimaryButton>
                  ) : (
                    <PrimaryButton
                      className={styles.primaryBtn}
                      onClick={() => {
                        addItemFunction();
                      }}
                    >
                      Save
                    </PrimaryButton>
                  )}
                  {props.currentItem !== undefined ? (
                    <DeleteItem
                      context={props.context}
                      setItems={props.setItems}
                      hidePopup={props.hidePopup}
                      currentItem={props.currentItem}
                      getItems={props.getItems}
                    />
                  ) : null}
                  <DefaultButton
                    onClick={() => {
                      props.hidePopup();
                      setData();
                    }}
                  >
                    Cancel
                  </DefaultButton>
                </div>
                {props.isRequestManager ? (
                  <DefaultButton
                    onClick={() => {
                      sendToDeliveryDepartment();
                    }}
                  >
                    Send to delivery department
                  </DefaultButton>
                ) : null}
              </div>
              
            </FocusTrapZone>
          </Popup>
        </Layer>
      )}
      <Dialog
        hidden={hideDialog}
        onDismiss={toggleHideDialog}
        dialogContentProps={dialogContentProps}
       >
        <DialogFooter className={styles.errorDialog}>
          <DefaultButton onClick={toggleHideDialog} text="Close" />
        </DialogFooter>
     </Dialog> 
    </>
  );
};
