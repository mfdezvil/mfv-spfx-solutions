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
                        </tr>
                    </thead>
                    <tbody>
                        {
                            userTeams.map((email) => {
                                return <tr>
                                    <th>{email.displayName}</th>
                                    <th>{email.description}</th> 
                                </tr>;
                            })
                        }
                    </tbody>
                </table>
            </div>
        );
    }

}