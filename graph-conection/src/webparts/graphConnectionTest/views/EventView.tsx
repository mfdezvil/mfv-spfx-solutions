import * as React from 'react';
import {IEventItem} from '../models/IEventItem';
import styles from '../components/GraphConnectionTest.module.scss';
import * as strings from 'GraphConnectionTestWebPartStrings';

export class EventView extends React.Component {

    
    public static GetEventsViewTable(userEvents: IEventItem[]):JSX.Element {
        return (
            <div>
                <table className={styles.table}>
                    <thead>
                        <tr>
                            <th>{strings.Organizer}</th>
                            <th>{strings.Subject}</th>
                            <th>{strings.Start}</th> 
                            <th>{strings.End}</th> 
                            <th>{strings.Location}</th> 
                        </tr>
                    </thead>
                    <tbody>
                        {
                            userEvents.map((event) => {
                                return <tr>
                                    <th>{event.organizer}</th>
                                    <th>{event.subject}</th> 
                                    <th>{event.start}</th>
                                    <th>{event.end}</th> 
                                    <th>{event.location}</th> 
                                </tr>;
                            })
                        }
                    </tbody>
                </table>
            </div>
        );
    }

}