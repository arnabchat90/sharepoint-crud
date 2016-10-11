import * as React from 'react';
import { css, Button } from 'office-ui-fabric-react';
import { HttpClient } from '@microsoft/sp-client-base';
import styles from '../SharepointCrud.module.scss';
import { ISharepointCrudWebPartProps } from '../ISharepointCrudWebPartProps';
import {IListItem} from "./SharepointCrud"

export interface IResultListProps {
  items : IListItem[];
}

export default class ResultList extends React.Component<IResultListProps,{}> {
  constructor(props: IResultListProps) {
    super(props);
  }

  render() : JSX.Element {
     const items : JSX.Element[] = this.props.items.map((item: IListItem, i: number) : JSX.Element => {
      return (
        <li>{item.Title} ({item.Id})</li>
      );
    });
    return (
      <div id="resultDiv" className={css('ms-Grid-row ms-bgColor-neutralSecondary ms-fontColor-white', styles.row)}>
        Following are the items in your list :
        <ul>
          {items}
        </ul>
      </div>
    );
  }
}