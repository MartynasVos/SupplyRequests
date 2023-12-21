import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IListWebPartWebPartProps } from "../ListWebPartWebPart";
import { List } from "./List";

export interface IDetailsListProps {
  context: WebPartContext;
}

export default class ListWebPart extends React.Component<
  IListWebPartWebPartProps,
  {}
> {
  public render(): React.ReactElement<IListWebPartWebPartProps> {
    const { context } = this.props;
    return (
      <section>
        <List context={context} />
      </section>
    );
  }
}
