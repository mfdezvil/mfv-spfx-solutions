import * as React from 'react';
import styles from './GraphConnectionTest.module.scss';
import * as strings from 'GraphConnectionTestWebPartStrings';
import { IGraphConnectionTestProps } from './IGraphConnectionTestProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { IGraphConnectionTestState } from './IGraphConnectionTestState';
import { GraphAccessManager } from '../managers/GraphAccessManager';
import {
  autobind,
  Spinner, SpinnerSize,
 } from 'office-ui-fabric-react';
import { GraphDataOptions } from '../utils/GraphDataOptions';
import { Logger } from '../utils/Logger';
import {MainView} from '../views/MainView';

export default class GraphConnectionTest extends React.Component<IGraphConnectionTestProps, IGraphConnectionTestState> {

  private graphAccessManager: GraphAccessManager;
  constructor(props: IGraphConnectionTestProps, state: IGraphConnectionTestState) {
    super(props);
    this.graphAccessManager = new GraphAccessManager(this.props.context.msGraphClientFactory, this.props.context.pageContext.cultureInfo.currentCultureName);
    this.state= {
      isLoading: false,
      currentUserData: null,
      userMails: null,
      userEvents: null,
      userTeams: null,
      errorMessage: ""
    };

  }

  public render(): React.ReactElement<IGraphConnectionTestProps> {
    let displayContent= this.GetDisplayTextBySelectedGraphDataOption();
    
    let loading= null;
    if(this.state.isLoading) {
      loading=<div><Spinner size={SpinnerSize.medium} /></div>;
    }
    let errorMessage= null;
    if(this.state.errorMessage != "") {
      errorMessage =<div>{this.state.errorMessage}</div>;
    }

    return (
      <div className={ styles.graphConnectionTest }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>{strings.WelcomeWebPartText}</span>
              <p className={ styles.subTitle }>{strings.WelcomeDescriptionText}</p>
            </div>
          </div>
          <div className={ styles.row }>
            <div className={ styles.column }>
              {displayContent}
              {loading}
              {errorMessage}
            </div>
          </div>
        </div>
      </div>
    );
  }

  @autobind
  private showUserData(): void {
    this.GetCurrentUserData();
  }

  @autobind
  private showMails(): void {
    this.GetUserMails();
  }

  @autobind
  private showEvents(): void {
    this.GetUserEvents();
  }

  @autobind
  private showTeams(): void {
    this.GetUserTeams();
  }

  private GetDisplayTextBySelectedGraphDataOption():JSX.Element {
    let displayContent= null;
    
    switch(this.props.selectedGraphDataOption) {
      case null: 
      case undefined:
        displayContent = <p className={styles.description}>{strings.NoDataOptionSelected}</p>;
        break;
      case GraphDataOptions.currentUser:
        displayContent = MainView.GetViewUserData(this.state, this.showUserData);
        break;
      case GraphDataOptions.mails:
        displayContent= MainView.GetViewMailData(this.state, this.showMails);
        break;
      case GraphDataOptions.calendar:
        displayContent= MainView.GetViewEventsData(this.state, this.showEvents);
        break;
      case GraphDataOptions.teams:
        displayContent= MainView.GetViewTeamsData(this.state, this.showTeams);
        break;
      
      default:
        displayContent = <p className={styles.description}>{strings.NotSupportedOption}</p>;
        Logger.WriteError("GetDisplayTextBySelectedGraphDataOption", "Not supported option in GraphDataOptions");
        break;
    }
    return displayContent;
  }

  private GetCurrentUserData():void {
    var _self= this;
    this.setState({...this.state, isLoading: true});
    try{
      this.graphAccessManager.GetUserData().then(userData => {
        this.setState({...this.state, isLoading: false});
        if(!userData) {
          _self.setState({..._self.state, errorMessage: strings.CannotLoadUserData});
          return;
        }
        _self.setState({..._self.state, currentUserData: userData, errorMessage: ""});

      });
    } 
    catch (error) {
      Logger.WriteError("GetCurrentUserData", error);
      this.setState({...this.state, isLoading: false, errorMessage: strings.CannotLoadUserData});
    }
  }

  private GetUserMails():void {
    var _self= this;
    this.setState({...this.state, isLoading: true});
    try{
      this.graphAccessManager.GetUserMails().then(userMails => {
        this.setState({...this.state, isLoading: false});
        if(!userMails) {
          _self.setState({..._self.state, errorMessage: strings.CannotLoadUserMails});
          return;
        }
        _self.setState({..._self.state, userMails: userMails, errorMessage: ""});

      });
    } 
    catch (error) {
      Logger.WriteError("GetUserMails", error);
      this.setState({...this.state, isLoading: false, errorMessage: strings.CannotLoadUserMails});
    }
  }

  private GetUserEvents():void {
    var _self= this;
    this.setState({...this.state, isLoading: true});
    try{
      this.graphAccessManager.GetUserEvents().then(userEvents => {
        this.setState({...this.state, isLoading: false});
        if(!userEvents) {
          _self.setState({..._self.state, errorMessage: strings.CannotLoadUserEvents});
          return;
        }
        _self.setState({..._self.state, userEvents: userEvents, errorMessage: ""});

      });
    } 
    catch (error) {
      Logger.WriteError("GetUserEvents", error);
      this.setState({...this.state, isLoading: false, errorMessage: strings.CannotLoadUserEvents});
    }
  }

  private GetUserTeams():void {
    var _self= this;
    this.setState({...this.state, isLoading: true});
    try{
      this.graphAccessManager.GetUserTeams().then(userTeams => {
        this.setState({...this.state, isLoading: false});
        if(!userTeams) {
          _self.setState({..._self.state, errorMessage: strings.CannotLoadUserTeams});
          return;
        }
        _self.setState({..._self.state, userTeams: userTeams, errorMessage: ""});

      });
    } 
    catch (error) {
      Logger.WriteError("GetUserTeams", error);
      this.setState({...this.state, isLoading: false, errorMessage: strings.CannotLoadUserTeams});
    }
  }

}
