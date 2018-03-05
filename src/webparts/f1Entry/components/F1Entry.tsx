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

import {
  Fabric,
  DefaultButton
} from 'office-ui-fabric-react';


import { ListService } from '../common/services/ListService';

export interface IF1RaceList {
  value: IF1Race[];
}
export interface IF1Race {
  Title: string;
  RaceDate: string;
}

export default class F1Entry extends React.Component<IF1EntryProps, any> {
  private LIST_TITLE_RACES: string = "F1_Races";
  private LIST_TITLE_ENTRIES: string = "F1_Entries";
  private _listService: ListService;

  constructor(props) {
    super(props);
    this.state = { showEntryForm: false };
    this._listService = new ListService(this.props.context.spHttpClient);
    this._showEntryForm = this._showEntryForm.bind(this);
  }

  public render(): React.ReactElement<IF1EntryProps> {
    return (
      <Fabric>
      <div className={styles.f1Entry}>
        <div className={styles.container}>
          <div className={styles.row} hidden={this.state.showEntryForm}>
            <div className={styles.column}>
              <span className={styles.title}>Welcome to F1 Oracle!</span>
              <p className={styles.description} hidden={!this.state.nextRaceTitle} >Next race is: {this.state.nextRaceTitle}</p>
              <br />
              <div hidden={!this.state.userEntryChecked || this.state.userHasEntry}>
                <p>You do not have an entry</p>
                <br />                
                  <DefaultButton onClick={this._showEntryForm}>Place Entry</DefaultButton>                
              </div>
              <div hidden={!this.state.userEntryChecked || !this.state.userHasEntry}>
                <p>You have already entered:</p>
              </div>
            </div>
          </div>
          <div className={styles.row} hidden={!this.state.showEntryForm}>
            <div className={styles.column}>
              <span className={styles.title}>Enter your predictions!</span>
              <br />
              <div>Form Goes here</div>
              <br />
              <p><button onClick={this._showEntryForm}>OK</button><button onClick={this._showEntryForm}>Cancel</button></p>
            </div>
          </div>
        </div>
      </div>
      </Fabric>
    );
  }

  public componentDidMount() {
    this._getNextRace()
      .then(raceResponse => {
        let nextRace: string = raceResponse.value[0].Title;
        this.setState(() => {
          return { nextRaceTitle: nextRace };
        });

        this._getUserEntry().then(entryResponse => {
          let hasEntry = entryResponse.value.length > 0;
          this.setState(() => {
            return {
              userHasEntry: hasEntry,
              userEntryChecked: true
            };
          });

        });


      });
  }


  private _getNextRace(): Promise<IF1RaceList> {
    let q: string = `<View><Query><Where><Geq><FieldRef Name='ID'/><Value Type='Number'>2</Value></Geq></Where></Query><RowLimit>100</RowLimit></View>`;

    return this._listService.getListItemsByQuery(this.props.context.pageContext.web.absoluteUrl, this.LIST_TITLE_RACES, q);
  }

  private _getUserEntry(): Promise<any> {
    let q: string = `
    <View><Query><Where><And><Eq><FieldRef Name='Race' /><Value Type='Lookup'>` + this.state.nextRaceTitle + `</Value></Eq>
    <Eq><FieldRef Name='Author' /><Value Type='User'>` + this.props.context.pageContext.user.displayName + `</Value></Eq></And></Where></Query></View>`;

    return this._listService.getListItemsByQuery(this.props.context.pageContext.web.absoluteUrl, this.LIST_TITLE_ENTRIES, q);
  }

  private _showEntryForm() {
    this.setState(prevState => ({
      showEntryForm: !prevState.showEntryForm
    }));
  }
}
