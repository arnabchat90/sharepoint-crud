import * as React from 'react';
import { css, Button } from 'office-ui-fabric-react';
import { HttpClient } from '@microsoft/sp-client-base';
import styles from '../SharepointCrud.module.scss';
import { ISharepointCrudWebPartProps } from '../ISharepointCrudWebPartProps';
import ResultList,{IResultListProps} from "./ResultList"

export interface ISharepointCrudProps extends ISharepointCrudWebPartProps {
    httpClient : HttpClient,
    siteUrl : string,
    listName : string
}

export interface ISharepointCrudState {
  status : string;
  showDiv : boolean;
  items : IListItem[];
}
export interface IListItem {
  Title? : string;
  Id : number;
}

export default class SharepointCrud extends React.Component<ISharepointCrudProps, ISharepointCrudState> {

  constructor(props: ISharepointCrudProps, state : ISharepointCrudState){
    super(props);
    this.state = {
      status : this.listNotConfigured(this.props) ? 'Please configure list in web part properties': 'Ready',
      showDiv : false,
      items : []
    }
  }

  public listNotConfigured(props : ISharepointCrudProps) : boolean {
    return props.listName === undefined || props.listName == null || props.listName.length == 0;
  }

  public render(): JSX.Element {
    const items : JSX.Element[] = this.state.items.map((item: IListItem, i: number) : JSX.Element => {
      return (
        <li>{item.Title} ({item.Id})</li>
      );
    });
     const resultList: React.ReactElement<IResultListProps> = React.createElement(ResultList, {
       items : this.state.items
    });

    return (
      <div className={styles.sharepointCrud}>
        <div className={styles.container}>
          <div className={css('ms-Grid-row ms-bgColor-themeDark ms-fontColor-white', styles.row)}>
            <div className='ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1'>
              <span className='ms-font-xl ms-fontColor-white'>
                Welcome to SharePoint!
              </span>
              <p className='ms-font-l ms-fontColor-white'>
                Customize SharePoint experiences using Web Parts.
              </p>
              <p className='ms-font-l ms-fontColor-white'>
                {this.props.description}
              </p>
              <p className='ms-font-l ms-fontColor-white'>
                Status : {this.state.status}
              </p>
              <p className='ms-font-l ms-fontColor-white'>
                {this.props.listName}
              </p>
              </div>
            <div className='ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1'>
              <Button disabled={this.listNotConfigured(this.props) } onClick={() => this.readItems() }>Read all items</Button>

            </div>
          </div>
          <div>
            {this.state.showDiv ? <ResultList items={this.state.items} /> : null}
          </div>
        </div>
      </div>
    );
  }
  public readItems() : void {
    this.setState({
      status : "Loading all items",
      showDiv : false,
      items : []
    });
    this.props.httpClient.get(`${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items?select=Title/Id`, {
      headers : {
        'Accept' : 'application/json;odata=nometadata',
        'odata-version' : ''
      }
    }).then((response : Response) : Promise<{value : IListItem[]}> => {
      return response.json();
    }).then((response : {value : IListItem[]}) : void => {
      this.setState({
        status : `Successfully loaded ${response.value.length} items`,
        showDiv : true,
        items : response.value
      });
    }, (error : any) : void => {
      this.setState({
          status: 'Loading all items failed with error: ' + error.message,
          showDiv : false,
          items: []
        });
    });
  }
}
