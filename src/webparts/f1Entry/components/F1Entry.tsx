import * as React from 'react';
import styles from './F1Entry.module.scss';
import { IF1EntryProps } from './IF1EntryProps';
import { escape } from '@microsoft/sp-lodash-subset';
import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions
} from '@microsoft/sp-http';
import {
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';

export interface IF1RaceList{
  value: IF1Race[];
}
export interface IF1Race {
  Title: string;
  RaceDate: string;
}

export default class F1Entry extends React.Component<IF1EntryProps, any> {
  private LIST_TITLE_RACES:string = "F1_Races";

  public render(): React.ReactElement<IF1EntryProps> {
    return (
      <div className={styles.f1Entry}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}>Welcome to F1 Oracle!</span>              
              <p className={styles.description}>Next race is: {escape(this.state ? this.state.description : "-loading-")}</p>
            </div>
          </div>
          <div id="spListContainer" />
        </div>
      </div>
    );
  }

  public componentDidMount(){
     this._getListData()
      .then((response) => {
        let desc: string = '';
          desc += response.value[0].Title;
        this.setState(() => {
          return {description: desc};
        });
      });
  }

  
  private _getListData(): Promise<IF1RaceList> {
    return this.props.context.spHttpClient.get(
      this.props.context.pageContext.web.absoluteUrl +
      `/_api/web/lists/GetByTitle('` + this.LIST_TITLE_RACES + `')/items?$top=1`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }
}
