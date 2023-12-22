import * as React from "react";
import { Dialog, DialogType, DialogFooter } from "@fluentui/react/lib/Dialog";
import { DefaultButton } from "@fluentui/react/lib/Button";
import { useBoolean } from "@fluentui/react-hooks";
import { IDeleteItemProps } from "./EditItem";
import { SPFx, spfi } from "@pnp/sp";

const dialogContentProps = {
  type: DialogType.normal,
  title: "Are you sure you want to delete this request?",
  closeButtonAriaLabel: "Close",
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
        style={{ backgroundColor: "#f00", color: '#fff' }}
        onClick={toggleHideDialog}
        text="Delete"
      />
      <Dialog
        hidden={hideDialog}
        onDismiss={toggleHideDialog}
        dialogContentProps={dialogContentProps}
      >
        <DialogFooter>
          <DefaultButton
            style={{ backgroundColor: "#f00", color: '#fff' }}
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
