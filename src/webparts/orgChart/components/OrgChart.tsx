import { IWebPartContext } from '@microsoft/sp-webpart-base';
import { Placeholder } from "@pnp/spfx-controls-react";
import { IPropertyFieldGroupOrPerson } from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react';
import * as React from 'react';
import * as ReactOrgChart from 'react-orgchart';
import { IPerson } from '../../../interfaces/IPerson';
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
  constructor(props) {
    super(props);
    this.state = {
      errorHandlerProperties: { error: false, errorMsg: "" },
      error: false,
      node: null
    };
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
    if (this.props.useGraphApi !== nextProps.useGraphApi) {
      if (nextProps.useGraphApi) {

        this.props.dataService.getDirectReportsForUserFromGraphAPI(nextProps.selectedGraphUser).then(
          (person: IPerson) => {
            this.setState({ node: person });
          });
      }
      else {
        this.props.dataService.getDirectReportsForUser(nextProps.selectedList, nextProps.selectedUser).then(
          (person: IPerson) => {
            this.setState({ node: person });
          });
      }
    }
  }

  public componentDidMount() {
    if (this.props.useGraphApi) {
      this.props.dataService.getDirectReportsForUserFromGraphAPI(this.props.selectedGraphUser).then(
        (person: IPerson) => {
          this.setState({ node: person });
        });
    } else {
      this.props.dataService.getDirectReportsForUser(this.props.selectedList, this.props.selectedUser).then(
        (person: IPerson) => {
          this.setState({ node: person });
        });
    }
  }

  public render(): React.ReactElement<IOrgChartProps> {

    const CustomOrgChartNodeComponent = ({ node }) => {
      return (
        <OrgChartNodeComponent node={node} styleIsSmall={this.props.styleIsSmall}></OrgChartNodeComponent>
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

