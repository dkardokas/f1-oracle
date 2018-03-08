import * as React from 'react';
import styles from './F1Entry.module.scss';
import { IF1EntryProps } from './IF1EntryProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import { Fabric, DefaultButton, PrimaryButton } from 'office-ui-fabric-react';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { Modal } from 'office-ui-fabric-react/lib/Modal';
import { Dropdown, IDropdown, DropdownMenuItemType, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';

import { ListService } from '../common/services/ListService';

export interface IF1RaceList {
  value: IF1Race[];
}
export interface IF1Race {
  Title: string;
  RaceDate: string;
  Id: number;
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
  private _selectedP1?: string;


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
                <p className={styles.description} hidden={!this.state.nextRaceTitle} >Next race is: {this.state.nextRaceTitle}
                <br />
                on {this.state.nextRaceDate}
                </p>
                
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

          <Modal
            isOpen={this.state.showEntryForm}
            // onDismiss={ this._closeModal }
            isBlocking={true}
            containerClassName={styles.modalContainer}
          >
            <Dropdown
              label='Winner'
              options={this.state.driversList}
              onChanged={(item) => {
                this.setState(() => {
                  return {
                    P1: item
                  }
                })
              }}
            />
            <Dropdown
              label='P2'
              options={this.state.driversList}
              onChanged={(item) => {
                this.setState(() => {
                  return {
                    P2: item
                  }
                })
              }}
            />
            <Dropdown
              label='P3'
              options={this.state.driversList}
              onChanged={(item) => {
                this.setState(() => {
                  return {
                    P3: item
                  }
                })
              }}
            />
            <Dropdown
              label='P4'
              options={this.state.driversList}
              onChanged={(item) => {
                this.setState(() => {
                  return {
                    P4: item
                  }
                })
              }}
            />
            <Dropdown
              label='P5'
              options={this.state.driversList}
              onChanged={(item) => {
                this.setState(() => {
                  return {
                    P5: item
                  }
                })
              }}
            />

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
          </Modal>
        </div>
      </Fabric>
    );
  }


  public componentDidMount() {
    this._getNextRace().then(raceResponse => {
      this.setState(() => {
        var raceDate = new Date(raceResponse.value[0].RaceDate);
        return {
          nextRaceTitle: raceResponse.value[0].Title,
          nextRaceId: raceResponse.value[0].Id,
          nextRaceDate: raceDate.toDateString()
        };
      });

      this._getUserEntry().then(entryResponse => {
        let hasEntry = entryResponse.value.length > 0;
        this.setState(() => {
          return {
            userHasEntry: hasEntry,
            userEntryChecked: true
          };
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
    });

    this._getDriversList().then(driversResponse => {
      let driverOptions = [];
      let previousTeam: string = '';
      driversResponse.value.forEach(driver => {
        if (driver.Team != previousTeam) {
          if (previousTeam.length > 0) {
            driverOptions.push({ text: '-', itemType: DropdownMenuItemType.Divider });
          }
          driverOptions.push({ text: driver.Team, itemType: DropdownMenuItemType.Header });
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
    let q: string = `<View><Query><Where><Eq><FieldRef Name='Submission_Open' /><Value Type='Boolean'>1</Value></Eq></Where><OrderBy><FieldRef Name='Date' Ascending='True' /></OrderBy></Query></View>`;

    return this._listService.getListItemsByQuery(this._webUrl, this.LIST_TITLE_RACES, q);
  }

  private _getUserEntry(): Promise<any> {
    let q: string = `
    <View><Query><Where><And><Eq><FieldRef Name='Race' /><Value Type='Lookup'>` + this.state.nextRaceTitle + `</Value></Eq>
    <Eq><FieldRef Name='Author' /><Value Type='User'>` + this.props.context.pageContext.user.displayName + `</Value></Eq></And></Where></Query></View>`;

    return this._listService.getListItemsByQuery(this._webUrl, this.LIST_TITLE_ENTRIES, q);
  }

  private _getDriversList(): Promise<any> {
    let q: string = `<View><Query><OrderBy><FieldRef Name='Team' Ascending='True' /></OrderBy></Query></View>`

    return this._listService.getListItemsByQuery(this._webUrl, this.LIST_TITLE_DRIVERS, q);
  }

  @autobind
  private _showEntryForm() {
    this.setState((prevState, props) => {
      return {
        showEntryForm: !prevState.showEntryForm
      }
    }
    );
  }

  @autobind
  private _submitEntry() {
    let selectedData = {
      Entry_P1Id: this.state.P1.key,
      Entry_P2Id: this.state.P2.key,
      Entry_P3Id: this.state.P3.key,
      Entry_P4Id: this.state.P4.key,
      Entry_P5Id: this.state.P5.key,
      RaceId: this.state.nextRaceId
    }
    this._listService.createItem(this._webUrl, this.LIST_TITLE_ENTRIES, selectedData);
    this.setState((prevState, props) => {
      return {
        showEntryForm: false
      }
    });
  }
}
