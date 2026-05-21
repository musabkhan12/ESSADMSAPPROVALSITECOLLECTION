import * as React from "react";
import * as ReactDom from "react-dom";
import { getSP } from "./loc/pnpjsConfig";
import Provider from "../../GlobalContext/provider";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import Mdr from "./components/Mdr";
import { IMdrProps } from "./components/IMdrProps";

export interface IMdrWebPartProps {
  context: any;
  siteUrl: string;
}

export default class MdrWebPart extends BaseClientSideWebPart<IMdrWebPartProps> {
  public render(): void {
    const element = React.createElement(
      Provider,
      null,
      React.createElement(Mdr, {
        context: this.context,
        siteUrl: this.context.pageContext.web.absoluteUrl,
      }),
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    await super.onInit();
    getSP(this.context);
  }
}
