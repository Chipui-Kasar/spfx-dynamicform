import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { type IPropertyPaneConfiguration } from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";

import * as strings from "DynamicFormWebPartStrings";
import DynamicForm from "./components/DynamicForm";
import { IDynamicFormProps } from "./components/IDynamicFormProps";
import {
  PropertyFieldListPicker,
  PropertyFieldListPickerOrderBy,
} from "@pnp/spfx-property-controls/lib/PropertyFieldListPicker";
import { sp } from "@pnp/sp/presets/all";
export interface IDynamicFormWebPartProps {
  description: string;
  listName: string;
  lists: string;
}

export default class DynamicFormWebPart extends BaseClientSideWebPart<IDynamicFormWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = "";

  public async render(): Promise<void> {
    const element: React.ReactElement<IDynamicFormProps> = React.createElement(
      DynamicForm,
      {
        context: this.context,
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        lists: this.properties.lists,
        listName: this.properties.listName,
        onConfigure: () => {
          this.onConfigure();
        },
      }
    );
    await this.getListName(this.properties.lists);
    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    sp.setup({
      spfxContext: this.context as any,
    });
    return this._getEnvironmentMessage().then((message) => {
      this._environmentMessage = message;
    });
  }
  private onConfigure = (): void => {
    this.context.propertyPane.open();
  };

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app
        .getContext()
        .then((context) => {
          let environmentMessage: string = "";
          switch (context.app.host.name) {
            case "Office": // running in Office
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentOffice
                : strings.AppOfficeEnvironment;
              break;
            case "Outlook": // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentOutlook
                : strings.AppOutlookEnvironment;
              break;
            case "Teams": // running in Teams
            case "TeamsModern":
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentTeams
                : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(
      this.context.isServedFromLocalhost
        ? strings.AppLocalEnvironmentSharePoint
        : strings.AppSharePointEnvironment
    );
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty(
        "--bodyText",
        semanticColors.bodyText || null
      );
      this.domElement.style.setProperty("--link", semanticColors.link || null);
      this.domElement.style.setProperty(
        "--linkHovered",
        semanticColors.linkHovered || null
      );
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected async onPropertyPaneFieldChanged(): Promise<void> {
    if (this.properties.lists) {
      await this.getListName(this.properties.lists);
    }
    this.context.propertyPane.refresh();
  }
  public async getListName(listID: string): Promise<any> {
    //get list name using guid
    if (!listID) return;
    const listName = await sp.web.lists
      .getById(listID)
      .select("EntityTypeName")();
    const formattedName = listName.EntityTypeName.replace(
      /List(?:Item|List)?$/,
      ""
    ).replace("_x0020_", " ");
    this.properties.listName = formattedName;
    return this.properties.listName;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: "Select a list to see the dynamic form",
          },
          groups: [
            {
              groupName: "Basic",
              groupFields: [
                PropertyFieldListPicker("lists", {
                  label: "Select a list",
                  selectedList: this.properties.lists,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  deferredValidationTime: 0,
                  key: "listPickerFieldId",
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
