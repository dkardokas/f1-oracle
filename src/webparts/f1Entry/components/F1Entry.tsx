import * as React from 'react';
import styles from './F1Entry.module.scss';
import { IF1EntryProps } from './IF1EntryProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import { Fabric, DefaultButton, PrimaryButton } from 'office-ui-fabric-react';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { Dropdown, IDropdown, DropdownMenuItemType, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';

import { ListService } from '../common/services/ListService';

export interface IF1RaceList {
  value: IF1Race[];
}
export interface IF1Race {
  Title: string;
  RaceDate: string;
}
export interface IF1Driver {
  Title: string;
  Team: string;
}

export default class F1Entry extends React.Component<IF1EntryProps, any> {
  private LIST_TITLE_RACES: string = "F1_Races";
  private LIST_TITLE_ENTRIES: string = "F1_Entries";
  private LIST_TITLE_DRIVERS: string = "F1_Drivers";
  private _listService: ListService;
  private _webUrl: string;


  constructor(props) {
    super(props);
    this.state = { showEntryForm: false };
    this._listService = new ListService(this.props.context.spHttpClient);    
    this._webUrl = this.props.context.pageContext.web.absoluteUrl;
  }

  public render(): React.ReactElement<IF1EntryProps> {
    return (
      <Fabric>
        <div className={styles.f1Entry}>
          <div className={styles.container}>
            <div className={styles.row}>
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
          </div>
        </div>
        <Panel
          isOpen={this.state.showEntryForm}
          type={PanelType.medium}
          headerText='Enter your predictions'
          onRenderFooterContent={this._onRenderFooterContent}
        >
          <Dropdown
            label='Winner'
            options={this.state.driversList}
          />
          <Dropdown
            label='P2'
            options={this.state.driversList}
          />
          <Dropdown
            label='P3'
            options={this.state.driversList}
          />
          <Dropdown
            label='P4'
            options={this.state.driversList}
          />
          <Dropdown
            label='P5'
            options={this.state.driversList}
          />
        </Panel>
      </Fabric>
    );
  }

  @autobind
  private _onRenderFooterContent(): JSX.Element {
    return (
      <div>
        <PrimaryButton
          onClick={this._submitEntry}
          style={{ 'marginRight': '8px' }}
        >
          Save
        </PrimaryButton>
        <DefaultButton
          onClick={this._showEntryForm}
        >
          Cancel
        </DefaultButton>
      </div>
    );
  }

  public componentDidMount() {
    this._getNextRace().then(raceResponse => {
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
      })
    });

    this._getDriversList().then(driversResponse => {
      let driverOptions = [];
      let previousTeam:string = '';
      driversResponse.value.forEach(driver => {
        if(driver.Team != previousTeam){
          if(previousTeam.length > 0){
            driverOptions.push({ text: '-', itemType: DropdownMenuItemType.Divider });
          }
          driverOptions.push({text: driver.Team, itemType: DropdownMenuItemType.Header });
        }
        previousTeam = driver.Team;

        driverOptions.push({
          key: driver.Id,
          text: driver.Title
        });
      });
      this.setState(() => {
        return { driversList: driverOptions };
      });

    });
  }


  private _getNextRace(): Promise<IF1RaceList> {
    let q: string = `<View><Query><Where><Geq><FieldRef Name='ID'/><Value Type='Number'>2</Value></Geq></Where></Query><RowLimit>100</RowLimit></View>`;

    return this._listService.getListItemsByQuery(this._webUrl, this.LIST_TITLE_RACES, q);
  }

  private _getUserEntry(): Promise<any> {
    let q: string = `
    <View><Query><Where><And><Eq><FieldRef Name='Race' /><Value Type='Lookup'>` + this.state.nextRaceTitle + `</Value></Eq>
    <Eq><FieldRef Name='Author' /><Value Type='User'>` + this.props.context.pageContext.user.displayName + `</Value></Eq></And></Where></Query></View>`

    return this._listService.getListItemsByQuery(this._webUrl, this.LIST_TITLE_ENTRIES, q);
  }

  private _getDriversList(): Promise<any> {
    let q: string = `<View><Query><OrderBy><FieldRef Name='Team' Ascending='True' /></OrderBy></Query></View>`

    return this._listService.getListItemsByQuery(this._webUrl, this.LIST_TITLE_DRIVERS, q);
  }

  @autobind
  private _showEntryForm() {
    this.setState(prevState => ({
      showEntryForm: !prevState.showEntryForm
    }));
  }

  @autobind
  private _submitEntry(){
    this._listService.createItem(this._webUrl, this.LIST_TITLE_ENTRIES);
    this._showEntryForm();
  }
}
