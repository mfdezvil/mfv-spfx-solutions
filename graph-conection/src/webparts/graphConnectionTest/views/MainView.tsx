import * as React from 'react';
import { IGraphConnectionTestState } from '../components/IGraphConnectionTestState';
import { Utils } from '../utils/Utils';
import {
    PrimaryButton,
} from 'office-ui-fabric-react';
import * as strings from 'GraphConnectionTestWebPartStrings';
import styles from '../components/GraphConnectionTest.module.scss';
import { MailView } from '../views/MailView';
import { TeamView } from '../views/TeamView';
import { EventView } from '../views/EventView';

export class MainView extends React.Component {

  public static GetViewUserData(state: IGraphConnectionTestState, onClickButtonFunction:()=>void): JSX.Element {
      let displayText= null;
  
      if(!Utils.IsNullOrEmpty(state.currentUserData) && Utils.IsNullOrEmpty(state.errorMessage)) {
        let email=(!state.currentUserData.mail && !state.currentUserData.userPrincipalName) ? <span>{strings.NotDefinedUserEmail}</span> : 
            (state.currentUserData.mail) ? <span>{strings.YourEmailIs} {state.currentUserData.mail}</span> : <span>{strings.YourEmailIs} {state.currentUserData.userPrincipalName}</span>;
  
        displayText= 
          <div className={ styles.row }>
            <div className={ styles.column }>
              <p className={styles.description}>{strings.Hi}, {state.currentUserData.displayName}. {email}</p>
            </div>
          </div>;
      }
      else {
        displayText= <p className={styles.description}>
            <PrimaryButton
              text= {strings.GetMyUserData}
              title= "ButtonGetCurrentUserData"
              onClick={onClickButtonFunction}
            />
        </p>;
      }
      return displayText;
  }

  public static GetViewMailData(state: IGraphConnectionTestState, onClickButtonFunction:()=>void): JSX.Element {
      let displayText= null;
      if(state.userMails != null && Utils.IsNullOrEmpty(state.errorMessage)) {
        if(state.userMails.length==0) {
          displayText = <div className={ styles.row }><div className={ styles.column }><p className={styles.description}>{strings.HaveNotEmails}</p></div></div>;
        }
        else {
          displayText = MailView.GetMailsViewTable(state.userMails);
        }
      }
      else {
        displayText = 
          <p className={styles.description}>
            <PrimaryButton
              text= {strings.GetMyMails}
              title= "ButtonGetUserMails"
              onClick={onClickButtonFunction}
            />
          </p>;
      }
      return displayText;
  }

  public static GetViewTeamsData(state: IGraphConnectionTestState, onClickButtonFunction:()=>void): JSX.Element {
    let displayText= null;
    if(state.userTeams != null  && Utils.IsNullOrEmpty(state.errorMessage)) {
      if(state.userTeams.length==0) {
        displayText = <div className={ styles.row }><div className={ styles.column }><p className={styles.description}>{strings.HaveNotTeams}</p></div></div>;
      }
      else {
        displayText = TeamView.GetMailsViewTable(state.userTeams);
      }
    }
    else {
      displayText = 
        <p className={styles.description}>
          <PrimaryButton
            text= {strings.GetMyTeams}
            title= "ButtonGetUserTeams"
            onClick={onClickButtonFunction}
          />
        </p>;
    }
    return displayText;
  }

  public static GetViewEventsData(state: IGraphConnectionTestState, onClickButtonFunction:()=>void): JSX.Element {
    let displayText= null;
    if(state.userEvents != null && Utils.IsNullOrEmpty(state.errorMessage)) {
      if(state.userEvents.length==0) {
        displayText = <div className={ styles.row }><div className={ styles.column }><p className={styles.description}>{strings.HaveNotEvents}</p></div></div>;
      }
      else {
        displayText = EventView.GetEventsViewTable(state.userEvents);
      }
    }
    else {
      displayText = 
        <p className={styles.description}>
          <PrimaryButton
            text= {strings.GetMyEvents}
            title= "ButtonGetUserEvents"
            onClick={onClickButtonFunction}
          />
        </p>;
    }
    return displayText;
  }
}