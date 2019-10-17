import * as React from 'react';
import styles from './TestWebpart.module.scss';
import { ITestWebpartProps } from './ITestWebpartProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { ITestWebpartState } from './ITestWebpartState';
import {SharepointListsService,SPListRequestItems, SPBasicColumn, LookupSPColumn, PeopleSPColumn, 
  Utils, IBaseListItem, SPListItemLimit, SPListItemOrder, SPOrderType, SPListItemFilter, SPFilterType, SPListLookupItemFilter, SPFilterOperator} from 'webcommons-library';
import { ITestListItem } from '../models/ITestListItem';

interface ISPTestListItem extends IBaseListItem {
  SingleLine: string;
  MultiLine: string;
  Lookup: {Title:string; SingleLine2:string};
  People: {Title:string};
}

export default class TestWebpart extends React.Component<ITestWebpartProps, ITestWebpartState> {

  private sharepointListService: SharepointListsService; 

  constructor(props: ITestWebpartProps) {
    super(props);
    this.state = {
      text: "Hola",
      items: []
    };
    
    this.sharepointListService = new SharepointListsService(this.props.context.pageContext.web.absoluteUrl);
    this.GetItemsFromSP();
  }

  private GetItemsFromSP(): void {
    let _self= this;

    let listRequest: SPListRequestItems = this.GetRequestTenant1();
    //let listRequest: SPListRequestItems = this.GetRequestTenant2();

    this.sharepointListService.ReadItems<ISPTestListItem>(listRequest).then((results) => {
        let items: ITestListItem[]= [];
        if(!Utils.IsNullOrEmpty(results)) {
          results.forEach(item => {
            items.push({
              Title: item.Title,
              SingleLine: item.SingleLine,
              MultiLine: item.MultiLine,
              Lookup: (!Utils.IsNullOrEmpty(item.Lookup)) ? item.Lookup.Title + " - "+ item.Lookup.SingleLine2 : null,
              People: (!Utils.IsNullOrEmpty(item.People)) ? item.People.Title : null
            });
          });
        }
        _self.setState({..._self.state, items: items });
    });
    
  }

  private GetRequestTenant1(): SPListRequestItems {
    let listRequest: SPListRequestItems = new SPListRequestItems("TestList");
    listRequest.FieldsToSelect = [
      new SPBasicColumn("Title"),
      new SPBasicColumn("SingleLine"),
      new SPBasicColumn("MultiLine"),
      new LookupSPColumn("Lookup",["Title","SingleLine2"]),
      new PeopleSPColumn("People", ["Title"])
    ];
    
    listRequest.Filters= [
      new SPListItemFilter("SingleLine", SPFilterType.eq, "Campo Texto"),
      new SPListLookupItemFilter("Lookup", "SingleLine2", SPFilterType.eq, "Hola", SPFilterOperator.or)
    ];
    return listRequest;
  }

  private GetRequestTenant2(): SPListRequestItems {
    let listRequest: SPListRequestItems = new SPListRequestItems("MetaDataListDirectory", null, "https://mfvcyc2.sharepoint.com/sites/HubCyCEN/Lists/MetadataListDirectory/" );
    listRequest.FieldsToSelect = [
      new SPBasicColumn("Title")
    ];
    listRequest.Order = new SPListItemOrder("ID", SPOrderType.Desc);
    listRequest.Limit= new SPListItemLimit(5);
    return listRequest;
  }

  public render(): React.ReactElement<ITestWebpartProps> {
    return (
      <div className={ styles.testWebpart }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
              <p className={ styles.description }>{escape(this.props.description)}</p>
              <p className={ styles.description }>Items from SP List: </p>
              <p className={ styles.description }>
                <table className={styles.table}>
                  <tr>
                      <th>Title Field</th>
                      <th>SingleLine Field</th> 
                      <th>MultiLine Field</th> 
                      <th>Lookup Field</th> 
                      <th>People Field</th> 
                  </tr>
                  {
                    this.state.items.map((item) => {
                      return <tr className={styles.tableContent}>
                        <th>{escape(item.Title)}</th>
                        <th>{escape(item.SingleLine)}</th>
                        <th>{escape(item.MultiLine)}</th>
                        <th>{escape(item.Lookup)}</th>
                        <th>{escape(item.People)}</th>
                      </tr>;
                    })
                  }
                </table>        
              </p>
              <a href="https://aka.ms/spfx" className={ styles.button }>
                <span className={ styles.label }>Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
