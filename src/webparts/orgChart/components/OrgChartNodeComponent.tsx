import { Callout, DirectionalHint, getInitials, IPersona, Persona, PersonaSize } from 'office-ui-fabric-react';
import * as React from 'react';
import { IPerson } from '../../../interfaces/IPerson';
import DataService from '../../../services/dataservice';
import styles from './OrgChartNodeComponent.module.scss';

export interface IOrgChartNodeComponentProps {
  node: IPerson;
  styleIsSmall: boolean;
  dataService: DataService;
}

export interface IOrgChartNodeComponentState {
  isCalloutVisible?: boolean;
  directionalHint?: DirectionalHint;
  isBeakVisible?: boolean;
  imageUrl: string;
  jobTitle: string;
  department: string;
}

export default class OrgChartNodeComponent extends React.Component<IOrgChartNodeComponentProps, IOrgChartNodeComponentState> {
  private _persona: IPersona;
  private _callOut: any;
  private _menuButtonElement: HTMLElement | null;

  constructor(props) {
    super(props);
    this._persona = {
      imageInitials: getInitials(this.props.node.name, false),
      text: this.props.node.name,
    };

    this.state = {
      isCalloutVisible: false,
      isBeakVisible: true,
      directionalHint: DirectionalHint.bottomAutoEdge,
      imageUrl: this.props.node.imageUrl || null,
      jobTitle: this.props.node.description || null,
      department: this.props.node.department || null,
    };
  }

  public componentDidMount() {
    if (this.props.node.email) {
      this.props.dataService.getUserPhotoFromGraphApi(this.props.node.email).then(
        (blob: any) => {
          this.setState({ imageUrl: window.URL.createObjectURL(blob) });
        }
      );

      this.props.dataService.getUserInfoFromGraphApi(this.props.node.email).then(
        (userInfo: any) => {
          this.setState({ jobTitle: userInfo.jobTitle, department: userInfo.department });
        }
      );
    }
  }

  private _onCalloutDismiss = (): void => {
    this.setState({
      isCalloutVisible: false
    });
  }

  private _onShowMenuClicked = (): void => {
    this.setState({
      isCalloutVisible: true
    });
  }

  public render(): React.ReactElement<IOrgChartNodeComponentProps> {
    const { isCalloutVisible, isBeakVisible, directionalHint } = this.state;

    return (
      <div className={this.props.styleIsSmall ? styles.nodeBase : styles.nodeBig} ref={(menuButton) => this._menuButtonElement = menuButton} onMouseEnter={this._onShowMenuClicked} onMouseLeave={this._onCalloutDismiss}>
        <Persona
          {...this._persona}
          size={this.props.styleIsSmall ? PersonaSize.size72 : PersonaSize.size48}
          hidePersonaDetails={this.props.styleIsSmall}
          imageUrl={this.state.imageUrl}
          secondaryText={this.state.department}
        />
        {(isCalloutVisible) ? (
          <Callout
            className='ms-CalloutExample-callout'
            target={this._menuButtonElement}
            isBeakVisible={isBeakVisible}
            onDismiss={this._onCalloutDismiss}
            directionalHint={directionalHint}>
            <div className='ms-CalloutExample-header'>
              <p className='ms-CalloutExample-title'>
                {this.props.node.name}
              </p>
            </div>
            {(this.state.jobTitle) ? (
              <div className='ms-CalloutExample-inner'>
                <div className='ms-CalloutExample-content'>
                  <p className='ms-CalloutExample-subText'>
                    {this.state.jobTitle}
                  </p>
                </div>
              </div>) : (null)}
          </Callout>
        ) : (null)}
      </div>
    );
  }
}