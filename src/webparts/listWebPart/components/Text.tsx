import * as React from "react";
import { Dialog, DialogType, DialogFooter } from "@fluentui/react/lib/Dialog";
import { PrimaryButton, DefaultButton } from "@fluentui/react/lib/Button";
import { useBoolean } from "@fluentui/react-hooks";
import { IDeleteItemProps } from "./EditItem";
import { SPFx, spfi } from "@pnp/sp";

const dialogContentProps = {
  type: DialogType.normal,
  title: "Missing Subject",
  closeButtonAriaLabel: "Close",
  subText: "Do you want to send this message without a subject?",
};

export const DeleteItem = (
  props: IDeleteItemProps
): React.ReactElement<unknown, React.JSXElementConstructor<unknown>> => {
  const [hideDialog, { toggle: toggleHideDialog }] = useBoolean(true);
  function deleteItemFunction(): void {
    const deleteItem = async (): Promise<void> => {
      const sp = spfi().using(SPFx(props.context));
      const list = sp.web.lists.getByTitle("Requests");
      const i = await list.items.getById(props.currentItem.Id).delete();
      console.log(i);
    };
    deleteItem().then(
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
  return (
    <>
      <DefaultButton
        style={{ backgroundColor: "#f00" }}
        onClick={toggleHideDialog}
        text="Delete"
      />
      <Dialog
        hidden={hideDialog}
        onDismiss={toggleHideDialog}
        dialogContentProps={dialogContentProps}
      >
        <DialogFooter>
          <PrimaryButton
            onClick={() => {
              toggleHideDialog();
              deleteItemFunction();
            }}
            text="Delete Request"
          />
          <DefaultButton onClick={toggleHideDialog} text="Cancel" />
        </DialogFooter>
      </Dialog>
    </>
  );
};
