import { Environment, EnvironmentType, Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, IPropertyPaneDropdownOption, PropertyPaneDropdown, PropertyPaneToggle } from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import pnp, { ProcessHttpClientResponseException } from "@pnp/pnpjs";
import { IPropertyFieldGroupOrPerson, PrincipalType, PropertyFieldPeoplePicker } from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';
import * as strings from 'OrgChartWebPartStrings';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { PropertyPaneCreateListDialog } from '../../controls/CreateListDialog/PropertyPaneCreateListDialog';
import { ErrorObjectFormat } from '../../helpers/ErrorHandler';
import { IDataService } from '../../interfaces/IDataService';
import { IList } from "../../interfaces/IList";
import { IPersonListItem } from "../../interfaces/IPersonListItem";
import DataService from '../../services/dataservice';
import MockDataService from '../../services/mockdataservice';
import OrgChart, { ErrorHandlerProps, IOrgChartProps } from './components/OrgChart';

export interface IOrgChartWebPartProps {
  selectedList: string;
  selectedUser: string;
  selectedStyleSmall: boolean;
  createConfigList: any;
  selectedGraphUser: IPropertyFieldGroupOrPerson;
  useGraphApi: boolean;
  dataService: DataService;
}

export default class OrgChartWebPart extends BaseClientSideWebPart<IOrgChartWebPartProps> {
  private loadingIndicator = false;
  private _errorProps: ErrorHandlerProps = { errorMsg: "", error: false };
  private _dataService: IDataService;

  private _listDropDownOptions: IPropertyPaneDropdownOption[] = [];
  private _userDropDownOptions: IPropertyPaneDropdownOption[] = [];

  protected onInit(): Promise<void> {
    pnp.setup({
      spfxContext: this.context
    });
    if (!this._dataService) {
      if (Environment.type in [EnvironmentType.Local, EnvironmentType.Test]) {
        this._dataService = new MockDataService();
      }
      else {
        this._dataService = new DataService(this.context);
      }
    }
    return Promise.resolve();
  }

  private _createConfigList(listName: string): Promise<IList> {
    return this._dataService.checkIfListAlreadyExists(listName).then((exists) => {
      if (exists) {
        return Promise.reject({ message: "List already exists." });
      } else {
        return this._dataService.createList(listName).then((result: IList) => {
          this._listDropDownOptions.push({ key: result.Id, text: result.Title });
          this.context.propertyPane.refresh();
          return result;
        }).catch((error) => {
          return Promise.reject(error);
        });
      }
    });
  }

  private _setErrorProps(error: ErrorObjectFormat | ProcessHttpClientResponseException) {
    this._errorProps = <ErrorHandlerProps>{ error: true, errorMsg: error.statusText };
  }

  private _resetErrorProps() {
    this._errorProps = <ErrorHandlerProps>{ error: false, errorMsg: '' };
  }

  public render(): void {
    const element: React.ReactElement<IOrgChartProps> = React.createElement(
      OrgChart,
      {
        context: this.context,
        styleIsSmall: this.properties.selectedStyleSmall,
        errorHandlerProperties: this._errorProps,
        error: this._errorProps.error,
        dataService: this._dataService,
        useGraphApi: this.properties.useGraphApi,
        selectedGraphUser: this.properties.selectedGraphUser,
        selectedList: this.properties.selectedList,
        selectedUser: this.properties.selectedUser
      }
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

    this._dataService.getOrgList().then(
      (orgLists: IList[]) => {
        this._listDropDownOptions = orgLists.map((list) => { return { key: list.Id, text: list.Title }; });
        this.context.propertyPane.refresh();
        if (this.properties.selectedList) {
          return this._dataService.getUsersFromList(this.properties.selectedList);
        }
        else {
          this.loadingIndicator = false;
        }
      })
      .then((persons: IPersonListItem[]) => {
        if (persons && persons.length > 0) {
          this._userDropDownOptions = persons.map((user: IPersonListItem) => { return { key: user.Id, text: user.Title }; });
        } else if (this.properties.selectedList) {
          this._userDropDownOptions = [];
          this._setErrorProps({ statusText: "No users configured in the selected Config List." });
        }
        this.loadingIndicator = false;
        this.context.propertyPane.refresh();
      }).catch((error: ErrorObjectFormat | ProcessHttpClientResponseException) => {
        this._setErrorProps(error);
      });
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    this._resetErrorProps();
    if (propertyPath === 'selectedList' && (newValue != oldValue)) {
      this.properties.selectedUser = undefined;
      this.context.propertyPane.refresh();
      this.loadingIndicator = true;

      this._dataService.getUsersFromList(this.properties.selectedList)
        .then((persons: IPersonListItem[]) => {
          if (persons && persons.length > 0) {
            this._userDropDownOptions = persons.map((user: IPersonListItem) => { return { key: user.Id, text: user.Title }; });
          }
          else {
            this._userDropDownOptions = [];
            this._setErrorProps({ statusText: "No users configured in the selected Config List." });
          }
          this.loadingIndicator = false;
          this.context.propertyPane.refresh();
        }).catch((error: ErrorObjectFormat | ProcessHttpClientResponseException) => {
          this._setErrorProps(error);
        });
    }
    if (propertyPath === 'selectedGraphUser' && newValue) {
      if (newValue && newValue.length > 0)
        this.properties.selectedGraphUser = newValue[0];
    }

    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    this.render();
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
                PropertyPaneToggle('useGraphApi', {
                  label: "Use AD data to build the org chart",
                  checked: false,
                }),
                PropertyPaneDropdown('selectedList', {
                  label: "Select Org Config List",
                  options: this._listDropDownOptions,
                  disabled: this.properties.useGraphApi
                }),
                PropertyPaneDropdown('selectedUser', {
                  label: "Select user to start building the Org-Chart from the config list",
                  options: this._userDropDownOptions,
                  disabled: (this._userDropDownOptions.length < 1 || this.properties.useGraphApi),
                  selectedKey: null
                }),
                PropertyFieldPeoplePicker('selectedGraphUser', {
                  label: 'Select user to start building the Org-Chart from AD data',
                  initialData: this.properties.selectedGraphUser ? [this.properties.selectedGraphUser] : null,
                  allowDuplicate: false,
                  principalType: [PrincipalType.Users],
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  context: this.context,
                  properties: this.properties,
                  onGetErrorMessage: null,
                  key: 'peopleFieldId',
                  multiSelect: false,
                  disabled: !this.properties.useGraphApi
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
