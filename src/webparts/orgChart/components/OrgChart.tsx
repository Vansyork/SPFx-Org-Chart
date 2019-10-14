import { IWebPartContext } from '@microsoft/sp-webpart-base';
import { Placeholder } from "@pnp/spfx-controls-react";
import { IPropertyFieldGroupOrPerson } from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react';
import * as React from 'react';
import * as ReactOrgChart from 'react-orgchart';
import { IPerson } from '../../../interfaces/IPerson';
import { Person } from '../../../models/person';
import DataService from '../../../services/dataservice';
import OrgChartNodeComponent from "../components/OrgChartNodeComponent";
import styles from './OrgChart.module.scss';

export interface IOrgChartState {
  errorHandlerProperties: ErrorHandlerProps;
  error: boolean;
  node: IPerson;
}

export interface IOrgChartProps {
  context: IWebPartContext;
  styleIsSmall: boolean;
  errorHandlerProperties: ErrorHandlerProps;
  error: boolean;
  useGraphApi: boolean;
  dataService: DataService;
  selectedGraphUser: IPropertyFieldGroupOrPerson;
  selectedList: string;
  selectedUser: string;
}

export interface ErrorHandlerProps {
  error: boolean;
  errorMsg: string;
}

export default class OrgChart extends React.Component<IOrgChartProps, IOrgChartState> {
  private _person: Person;
  constructor(props) {
    super(props);
    this.state = {
      errorHandlerProperties: { error: false, errorMsg: "" },
      error: false,
      node: null
    };
  }

  public setPersonSate() {
    this.setState({ node: this._person });
  }

  private _onConfigure(ctx: IWebPartContext) {
    // Context of the web part
    ctx.propertyPane.open();
  }

  private _removeMessageBar = (): void => {
    this.setState({ errorHandlerProperties: { errorMsg: "", error: false } });
    this.setState({ error: false });
  }

  public componentWillReceiveProps(nextProps: IOrgChartProps) {
    if (this.props.error !== nextProps.error) {
      this.setState({ error: nextProps.error });
    }
    if (this.props.errorHandlerProperties !== nextProps.errorHandlerProperties) {
      this.setState({ errorHandlerProperties: nextProps.errorHandlerProperties });
    }

    if (this.state.node) {
      this.setState({ node: null });
    }
    if (nextProps.useGraphApi && nextProps.selectedGraphUser) {
      this._person = new Person(
        {
          Id: nextProps.selectedGraphUser.id,
          Title: nextProps.selectedGraphUser.fullName,
          email: nextProps.selectedGraphUser.email
        }, null, this.props.dataService, this.setPersonSate.bind(this));
    }
    else if (!nextProps.useGraphApi) {
      this.props.dataService.getDirectReportsForUser(nextProps.selectedList, nextProps.selectedUser).then(
        (person: IPerson) => {
          this.setState({ node: person });
        });
    }

  }

  public componentDidMount() {
    if (this.state.node) {
      this.setState({ node: null });
    }
    if (this.props.useGraphApi && this.props.selectedGraphUser) {
      this._person = new Person(
        {
          Id: this.props.selectedGraphUser.id,
          Title: this.props.selectedGraphUser.fullName,
          email: this.props.selectedGraphUser.email
        }, null, this.props.dataService, this.setPersonSate.bind(this));
    } else if (!this.props.useGraphApi && this.props.selectedList && this.props.selectedUser) {
      this.props.dataService.getDirectReportsForUser(this.props.selectedList, this.props.selectedUser).then(
        (person: IPerson) => {
          this.setState({ node: person });
        });
    }
  }

  public render(): React.ReactElement<IOrgChartProps> {

    const CustomOrgChartNodeComponent = ({ node }) => {
      return (
        <OrgChartNodeComponent node={node} styleIsSmall={this.props.styleIsSmall} dataService={this.props.dataService}></OrgChartNodeComponent>
      );
    };

    return (
      <div>
        {
          this.state.error ? (
            <MessageBar
              messageBarType={MessageBarType.error}
              isMultiline={false}
              onDismiss={this._removeMessageBar}
              dismissButtonAriaLabel='Close'>
              {this.state.errorHandlerProperties.errorMsg}
            </MessageBar>) : (null)
        }
        {
          this.state.node ? (
            <div className={styles.orgChart}>
              <div className={styles.container}>
                <ReactOrgChart tree={this.state.node} NodeComponent={CustomOrgChartNodeComponent} />
              </div>
            </div>
          ) : (
              <Placeholder
                iconName='Edit'
                iconText='Configure your web part'
                description='Please configure the web part.'
                buttonLabel='Configure'
                onConfigure={() => this._onConfigure(this.props.context)} />
            )
        }
      </div >
    );
  }
}

