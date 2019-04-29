import * as React from 'react';
import {ITeamItem} from '../models/ITeamItem';
import styles from '../components/GraphConnectionTest.module.scss';
import * as strings from 'GraphConnectionTestWebPartStrings';

export class TeamView extends React.Component {

    
    public static GetMailsViewTable(userTeams: ITeamItem[]):JSX.Element {
        return (
            <div>
                <table className={styles.table}>
                    <thead>
                        <tr>
                            <th>{strings.TeamName}</th>
                            <th>{strings.TeamDescription}</th> 
                            <th>{strings.TeamChannels}</th> 
                        </tr>
                    </thead>
                    <tbody>
                        {
                            userTeams.map((team) => {
                                return <tr>
                                    <th>{team.displayName}</th>
                                    <th>{team.description}</th> 
                                    <th className={styles.teamChannels}>
                                        <ul>
                                            {team.channels.map((channel) => {
                                                return <li>{channel.displayName}</li>;
                                            })}
                                        </ul>
                                    </th> 
                                </tr>;
                            })
                        }
                    </tbody>
                </table>
            </div>
        );
    }

}