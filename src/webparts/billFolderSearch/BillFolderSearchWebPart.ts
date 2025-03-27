import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  BaseClientSideWebPart,
  WebPartContext,
} from "@microsoft/sp-webpart-base";
import BillFolderSearch from "./components/BillFolderSearch";

export interface IBillFolderSearchProps {
  context: WebPartContext; // Pass the SPFx context
}

export default class BillFolderSearchWebPart extends BaseClientSideWebPart<{}> {
  public render(): void {
    const element: React.ReactElement<IBillFolderSearchProps> = React.createElement(
      BillFolderSearch,
      {
        context: this.context, // Pass context here
      }
    );
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }
}
