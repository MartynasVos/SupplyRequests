import * as React from "react";
import { SPFx, spfi } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import "@pnp/sp/taxonomy";
import { ITermInfo } from "@pnp/sp/taxonomy";
import { ISiteUserInfo } from "@pnp/sp/site-users/types";
import { IDetailsListProps } from "./ListWebPart";
import { ListItems } from "./ListItems";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ISiteGroupInfo } from "@pnp/sp/site-groups/types";
import { FormModalBox } from "./FormModalBox";
import { IFieldInfo } from "@pnp/sp/fields";
import "@pnp/sp/site-groups/web";
import { IDropdownOption } from "@fluentui/react/lib/Dropdown";
import type { IComboBoxOption } from "@fluentui/react";

export interface IListItemsProps {
  context: WebPartContext;
  items: IRequest[];
  requestTypes: IDropdownOption[];
  users: ISiteUserInfo[];
  isRequestManager: boolean;
  requestManagers: IComboBoxOption[] | undefined;
  taxonomy: IDropdownOption[];
  requestAreaChoices: IDropdownOption[] | undefined;
  setItems: React.Dispatch<React.SetStateAction<IRequest[]>>;
  getItems: () => Promise<IRequest[]>;
}
export interface IFormProps {
  context: WebPartContext;
  items: IRequest[];
  requestTypes: IDropdownOption[];
  users: ISiteUserInfo[];
  isRequestManager: boolean;
  taxonomy: IDropdownOption[];
  requestAreaChoices: IDropdownOption[] | undefined;
  setItems: React.Dispatch<React.SetStateAction<IRequest[]>>;
  getItems: () => Promise<IRequest[]>;
}
export interface IRequest {
  Id: number;
  Title: string;
  AuthorId: number;
  Description: string;
  RequestArea: string;
  RequestTypeId: number;
  Tags: object[];
  DueDate: Date;
  ExecutionDate: Date;
  Assigned_x0020_ManagerId: number;
  Status: string;
}
export interface IRequestTypes {
  Id: number;
  Title: string;
}

export const List = (
  props: IDetailsListProps
): React.ReactElement<unknown, React.JSXElementConstructor<unknown>> => {
  const [itemsState, setItems] = React.useState<IRequest[]>([]);
  const [requestTypes, setRequestTypes] = React.useState<IDropdownOption[]>([]);
  const [users, setUsers] = React.useState<ISiteUserInfo[]>([]);
  const [isRequestManager, setIsRequestManager] = React.useState(false);
  const [requestManagers, setRequestManagers] =
    React.useState<IComboBoxOption[]>();
  const [taxonomy, setTaxonomy] = React.useState<IDropdownOption[]>([]);
  const [requestAreaChoices, setRequestAreaChoices] =
    React.useState<IDropdownOption[]>();

  const sp = spfi().using(SPFx(props.context));
  const getItems = async (): Promise<IRequest[]> => {
    const items = await sp.web.lists.getByTitle("Requests").items();
    return items;
  };
  const getRequestTypes = async (): Promise<IRequestTypes[]> => {
    const requestTypes = await sp.web.lists
      .getByTitle("Request type")
      .items();
    return requestTypes;
  };
  const getUsers = async (): Promise<ISiteUserInfo[]> => {
    const users = await sp.web.siteUsers();
    return users;
  };
  const getUserGroup = async (): Promise<ISiteGroupInfo[]> => {
    const userGroup = await sp.web.currentUser.groups();
    return userGroup;
  };
  const getRequestManagers = async (): Promise<ISiteUserInfo[]> => {
    const users = await sp.web.siteGroups.getById(12).users();
    return users;
  };
  const getTaxonomy = async (): Promise<ITermInfo[]> => {
    const info: ITermInfo[] = await sp.termStore.groups
      .getById("57cb87c2-f752-4c56-8d61-dbe357db2d81")
      .sets.getById("d9e481e9-4309-4c4f-bd3a-588fc993ddc0")
      .terms();
    return info;
  };
  const getChoiceField = async (): Promise<IFieldInfo[]> => {
    const choiceField = await sp.web.lists
      .getByTitle("Requests")
      .fields.filter("Title eq 'Request Area'")
      .select("Choices")();
    return choiceField;
  };
  React.useEffect(() => {
    getItems().then(
      (result) => {
        setItems(result);
      },
      () => {
        return;
      }
    );
    getRequestTypes().then(
      (result) => {
        const arr: IDropdownOption[] = [];
        result.map((type) => {
          arr.push({ key: type.Id, text: type.Title });
        });
        setRequestTypes(arr);
      },
      () => {
        return;
      }
    );
    getUsers().then(
      (result) => {
        setUsers(result);
      },
      () => {
        return;
      }
    );
    getUserGroup().then(
      (result) => {
        result.map((group) => {
          if (group.Title === "Request Managers") {
            setIsRequestManager(true);
          }
        });
      },
      () => {
        return;
      }
    );
    getRequestManagers().then(
      (result) => {
        const arr: IComboBoxOption[] = [{ key: 0, text: '' }];
        result.map((manager) => {
          arr.push({ key: manager.Id, text: manager.Title });
        });
        setRequestManagers(arr);
      },
      () => {
        return;
      }
    );
    getTaxonomy().then(
      (result) => {
        const arr: IDropdownOption[] = [];
        result.map((tag) => {
          arr.push({ key: tag.id, text: tag.labels[0].name });
        });
        setTaxonomy(arr);
      },
      () => {
        return;
      }
    );
    getChoiceField().then(
      (result) => {
        const arr: IDropdownOption[] = [];
        result[0].Choices?.map((choice) => {
          arr.push({ key: choice, text: choice });
        });
        setRequestAreaChoices(arr);
      },
      () => {
        return;
      }
    );
  }, []);
  return (
    <div>
      <FormModalBox
        context={props.context}
        items={itemsState}
        requestTypes={requestTypes}
        users={users}
        isRequestManager={isRequestManager}
        taxonomy={taxonomy}
        requestAreaChoices={requestAreaChoices}
        setItems={setItems}
        getItems={getItems}
      />
      <ListItems
        context={props.context}
        items={itemsState}
        requestTypes={requestTypes}
        users={users}
        isRequestManager={isRequestManager}
        requestManagers={requestManagers}
        taxonomy={taxonomy}
        requestAreaChoices={requestAreaChoices}
        setItems={setItems}
        getItems={getItems}
      />
    </div>
  );
};
