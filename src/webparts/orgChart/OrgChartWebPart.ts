import { Environment, EnvironmentType, Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, IPropertyPaneDropdownOption, PropertyPaneDropdown, PropertyPaneToggle } from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import pnp, { ProcessHttpClientResponseException } from "@pnp/pnpjs";
import * as strings from 'OrgChartWebPartStrings';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { PropertyPaneCreateListDialog } from '../../controls/CreateListDialog/PropertyPaneCreateListDialog';
import { ErrorObjectFormat } from '../../helpers/ErrorHandler';
import { IDataService } from '../../interfaces/IDataService';
import { IList } from "../../interfaces/IList";
import { IPerson } from '../../interfaces/IPerson';
import { IPersonListItem } from "../../interfaces/IPersonListItem";
import DataService from '../../services/dataservice';
import MockDataService from '../../services/mockdataservice';
import OrgChart, { ErrorHandlerProps, IOrgChartProps } from './components/OrgChart';

export interface IOrgChartWebPartProps {
  selectedList: string;
  selectedUser: string;
  selectedStyleSmall: boolean;
  createConfigList: any;
}

export default class OrgChartWebPart extends BaseClientSideWebPart<IOrgChartWebPartProps> {
  private loadingIndicator = false;
  private _errorProps: ErrorHandlerProps = { errorMsg: "", error: false };
  private _dataService: IDataService;
  private get DataService(): IDataService {
    if (!this._dataService) {
      if (Environment.type in [EnvironmentType.Local, EnvironmentType.Test]) {
        this._dataService = new MockDataService();
      }
      else {
        this._dataService = new DataService(this.context);
      }
    }
    return this._dataService;
  }

  private _personNode: IPerson = null;
  private _listDropDownOptions: IPropertyPaneDropdownOption[] = [];
  private _userDropDownOptions: IPropertyPaneDropdownOption[] = [];
  protected onInit(): Promise<void> {
    pnp.setup({
      spfxContext: this.context
    });

    if (this.properties.selectedUser && this.properties.selectedUser) {
      return this.DataService.getDirectReportsForUser(this.properties.selectedList, this.properties.selectedUser).then(
        (person: IPerson) => {
          this._personNode = person;
          return Promise.resolve();
        })
        .catch((error) => {
          return Promise.reject(error);
        });
    } else {
      return Promise.resolve();
    }
  }

  private _createConfigList(listName: string): Promise<IList> {
    return this.DataService.checkIfListAlreadyExists(listName).then((exists) => {
      if (exists) {
        return Promise.reject({ message: "List already exists." });
      } else {
        return this.DataService.createList(listName).then((result: IList) => {
          this._listDropDownOptions.push({ key: result.Id, text: result.Title });
          this.context.propertyPane.refresh();
          return result;
        }).catch((error) => {
          return Promise.reject(error);
        });
      }
    })
  }

  private _setErrorProps(error: ErrorObjectFormat | ProcessHttpClientResponseException) {
    this._errorProps = <ErrorHandlerProps>{ error: true, errorMsg: error.statusText };
  }

  private _resetErrorProps() {
    this._errorProps = <ErrorHandlerProps>{ error: false, errorMsg: '' };
  }

  public render(): void {
    const element: React.ReactElement<IOrgChartProps> = React.createElement(
      OrgChart, //base react component
      {
        node: this._personNode,
        context: this.context,
        styleIsSmall: this.properties.selectedStyleSmall,
        errorHandlerProperties: this._errorProps,
        error: this._errorProps.error
      } // react properties
    );

    ReactDom.render(element, this.domElement); //inject into webpart dom
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onPropertyPaneConfigurationStart(): void {

    if (this._listDropDownOptions.length > 0) {
      return;
    }

    this.loadingIndicator = true;

    this.DataService.getOrgList().then(
      (orgLists: IList[]) => {
        this._listDropDownOptions = orgLists.map((list) => { return { key: list.Id, text: list.Title }; });
        this.context.propertyPane.refresh();
        if (this.properties.selectedList) {
          return this.DataService.getUsersFromList(this.properties.selectedList);
        }
        else {
          // clear status indicator
          this.loadingIndicator = false;
          // re-render the web part as clearing the loading indicator removes the web part body
          // this.render();
        }
      })
      .then((persons: IPersonListItem[]) => {
        if (persons && persons.length > 0) {
          this._userDropDownOptions = persons.map((user: IPersonListItem) => { return { key: user.Id, text: user.Title }; });
        } else if (this.properties.selectedList) {
          this._userDropDownOptions = [];
          this._setErrorProps({ statusText: "No users configured in the selected Config List." })
        }
        // clear status indicator
        this.loadingIndicator = false;
        // re-render the web part as clearing the loading indicator removes the web part body
        // this.render();
        // refresh the item selector control by repainting the property pane
        this.context.propertyPane.refresh();
      }).catch((error: ErrorObjectFormat | ProcessHttpClientResponseException) => {
        this._setErrorProps(error);
      });
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    this._resetErrorProps();
    if (propertyPath === 'selectedList' && newValue) {
      // push new list value
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      // reset selected item
      this.properties.selectedUser = undefined;
      this._personNode = null;
      // refresh the item selector control by repainting the property pane
      this.context.propertyPane.refresh();
      // communicate loading items
      this.loadingIndicator = true;

      this.DataService.getUsersFromList(this.properties.selectedList)
        .then((persons: IPersonListItem[]) => {
          if (persons && persons.length > 0) {
            this._userDropDownOptions = persons.map((user: IPersonListItem) => { return { key: user.Id, text: user.Title }; });
          }
          else {
            this._userDropDownOptions = [];
            this._setErrorProps({ statusText: "No users configured in the selected Config List." })
          }
          // clear status indicator
          this.loadingIndicator = false;
          // re-render the web part as clearing the loading indicator removes the web part body
          this.render();
          // refresh the item selector control by repainting the property pane
          this.context.propertyPane.refresh();
        });
    }
    if (propertyPath === 'selectedUser' && newValue) {
      if (this.properties.selectedUser && this.properties.selectedList) {
        // push new list value
        super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
        // reset selected item
        this._personNode = null;
        // communicate loading items

        this.DataService.getDirectReportsForUser(this.properties.selectedList, this.properties.selectedUser).then(
          (person: IPerson) => {
            this._personNode = person;
            // // re-render the web part as clearing the loading indicator removes the web part body
            this.render();
          })
          .catch((error: ErrorObjectFormat | ProcessHttpClientResponseException) => {
            this._setErrorProps(error);
          });
      }
      else {
        super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      }
    }
    else {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      showLoadingIndicator: this.loadingIndicator,    
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneDropdown('selectedList', {
                  label: "Select Org Config List",
                  options: this._listDropDownOptions
                }),
                PropertyPaneDropdown('selectedUser', {
                  label: "Select user to start building the Org-Chart",
                  options: this._userDropDownOptions,
                  disabled: (this._userDropDownOptions.length < 1),
                  selectedKey: null
                })
              ]
            },
            {
              groupName: "Style",
              groupFields: [
                PropertyPaneToggle('selectedStyleSmall', {
                  label: "Use small tiles",
                  checked: false
                })
              ]
            },
            {
              groupName: "Configuration Lists",
              groupFields: [
                new PropertyPaneCreateListDialog('createConfigList', {
                  buttonLabel: "Create Configuration List",
                  dialogTitle: "Create List",
                  dialogText: "Create a new org-chart configuration list with this dialog.",
                  saveAction: this._createConfigList.bind(this),
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
