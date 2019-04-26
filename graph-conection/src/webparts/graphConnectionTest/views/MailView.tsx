import * as React from 'react';
import {IMailItem} from '../models/IMailItem';
import styles from '../components/GraphConnectionTest.module.scss';
import * as strings from 'GraphConnectionTestWebPartStrings';

export class MailView extends React.Component {

    
    public static GetMailsViewTable(userMails: IMailItem[]):JSX.Element {
        return (
            <div>
                <table className={styles.table}>
                    <thead>
                        <tr>
                            <th>{strings.Received}</th>
                            <th>{strings.From}</th> 
                            <th>{strings.Subject}</th>
                            <th>{strings.Read}</th>
                        </tr>
                    </thead>
                    <tbody>
                        {
                            userMails.map((email) => {
                                return <tr>
                                    <th>{email.received}</th>
                                    <th>{email.from}</th> 
                                    <th>{email.subject}</th>
                                    <th>{email.isRead}</th>
                                </tr>;
                            })
                        }
                    </tbody>
                </table>
            </div>
        );
    }

}