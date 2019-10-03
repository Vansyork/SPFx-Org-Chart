import { IWebPartContext } from '@microsoft/sp-webpart-base';
import { Placeholder } from "@pnp/spfx-controls-react";
import { MessageBar, MessageBarType } from 'office-ui-fabric-react';
import * as React from 'react';
// import * as ReactOrgChart from './react-orgchart';
// import './react-orgchart/index.css';
import * as ReactOrgChart from 'react-orgchart';
import { IPerson } from '../../../interfaces/IPerson';
import OrgChartNodeComponent from "../components/OrgChartNodeComponent";
import styles from './OrgChart.module.scss';


export interface IOrgChartState {
  errorHandlerProperties: ErrorHandlerProps;
  error: boolean;
}

export interface IOrgChartProps {
  node: IPerson;
  context: IWebPartContext;
  styleIsSmall: boolean;
  errorHandlerProperties: ErrorHandlerProps;
  error: boolean;
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
      error: false
    };
  }

  private _onConfigure(ctx: IWebPartContext) {
    // Context of the web part
    ctx.propertyPane.open();
  }

  private _removeMessageBar = (): void => {
    this.setState({ errorHandlerProperties: { errorMsg: "", error: false } });
    this.setState({ error: false});
  }

  public componentWillReceiveProps(nextProps) {
    if (this.props.error !== nextProps.error) {
      this.setState({ error: nextProps.error });
    }
    if (this.props.errorHandlerProperties !== nextProps.errorHandlerProperties) {
      this.setState({ errorHandlerProperties: nextProps.errorHandlerProperties });
    }
  }

  public render(): React.ReactElement<IOrgChartProps> {

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
          this.props.node ? (
            <div className={styles.orgChart}>
              <div className={styles.container}>
                <ReactOrgChart tree={this.props.node} NodeComponent={OrgChartNodeComponent} isSmall={this.props.styleIsSmall} />
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

