import * as React from 'react';
import styles from './F1Entry.module.scss';
import { IF1EntryProps } from './IF1EntryProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import { Fabric, DefaultButton, PrimaryButton } from 'office-ui-fabric-react';
//import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { Modal } from 'office-ui-fabric-react/lib/Modal';
import { Dialog, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
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
    this.state = { showEntryForm: false, allselectedKeys: [] };
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
                  <p>1. {this.state.userEntry_P1}</p>
                  <p>2. {this.state.userEntry_P2}</p>
                  <p>3. {this.state.userEntry_P3}</p>
                  <p>4. {this.state.userEntry_P4}</p>
                  <p>5. {this.state.userEntry_P5}</p>
                  <br />
                  <p><DefaultButton onClick={this._showDeleteDialog}>Delete My Entry</DefaultButton></p>
                  <Dialog isOpen={this.state.showDeleteDialog}
                    title="Confirm Delete"
                    subText="Are you sure you want to remove your entry?"
                    onDismiss={this._showDeleteDialog}
                  >
                    <DialogFooter>
                      <PrimaryButton onClick={this._removeEntry} text='Delete' />
                      <DefaultButton onClick={this._showDeleteDialog} text='Cancel' />
                    </DialogFooter>

                  </Dialog>
                </div>
              </div>
            </div>
            <div id="results" className={styles.csstransforms}>
            <h2>Current Standings</h2>
              <table className={styles["table-header-rotated"]}>
                <thead>
                  <tr>
                    <th></th>
                    <th className={styles.rotate}><div><span>Australian GP</span></div></th>
                    <th className={styles.rotate}><div><span>Bahrain GP</span></div></th>
                    <th className={styles.rotate}><div><span>Chinese GP</span></div></th>
                    <th className={styles.rotate}><div><span>Total</span></div></th>
                    <th className={styles.rotate}><div><span>Best x?</span></div></th>
                  </tr>
                </thead>
                <tbody>
                  <tr>
                    <th className="row-header">Player 1</th>
                    <td>22</td>
                    <td>78</td>
                    <td>40</td>
                    <td>140</td>
                    <td>?</td>

                  </tr>
                  <tr>
                    <th className="row-header">Player 2</th>
                    <td>33</td>
                    <td>66</td>
                    <td>12</td>
                    <td>111</td>
                    <td>?</td>
                  </tr>
                  <tr>
                    <th className="row-header">Payer 3</th>
                    <td>44</td>
                    <td>56</td>
                    <td>0</td>
                    <td>100</td>
                    <td>?</td>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>

          <Modal
            isOpen={this.state.showEntryForm}
            isBlocking={true}
            containerClassName={styles.modalContainer}
          >
            <Dropdown
              label='Winner'
              options={this.state.driversList}
              onChanged={(item) => this._driverSelected('P1', item)}
            />
            <Dropdown
              label='P2'
              options={this.state.driversList}
              onChanged={(item) => this._driverSelected('P2', item)}

            />
            <Dropdown
              label='P3'
              options={this.state.driversList}
              onChanged={(item) => this._driverSelected('P3', item)}
            />
            <Dropdown
              label='P4'
              options={this.state.driversList}
              onChanged={(item) => this._driverSelected('P4', item)}
            />
            <Dropdown
              label='P5'
              options={this.state.driversList}
              onChanged={(item) => this._driverSelected('P5', item)}
            />

            <div>
              <p>{this.state.validationError}</p>
              <PrimaryButton
                onClick={this._submitEntry}
                style={{ 'marginRight': '8px' }}
                disabled={!this.state.selectionsValid}
              >
                Submit
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

      this._checkEntry();
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
    //let expandFields: Array<string> = ["Entry_P1", "Entry_P2", "Entry_P3", "Entry_P4", "Entry_P5"];
    let q: string = `
    <View><Query><Where><And><Eq><FieldRef Name='Race' /><Value Type='Lookup'>` + this.state.nextRaceTitle + `</Value></Eq>
    <Eq><FieldRef Name='Author' /><Value Type='User'>` + this.props.context.pageContext.user.displayName + `</Value></Eq></And></Where></Query><ViewFields><FieldRef Name='Entry_P1' /><FieldRef Name='Entry_P2' /><FieldRef Name='Entry_P3' /><FieldRef Name='Entry_P4' /><FieldRef Name='Entry_P5' /></ViewFields></View>`;

    return this._listService.getListItemsByQuery(this._webUrl, this.LIST_TITLE_ENTRIES, q);
  }

  private _getDriversList(): Promise<any> {
    let q: string = `<View><Query><OrderBy><FieldRef Name='Team' Ascending='True' /></OrderBy></Query></View>`;

    return this._listService.getListItemsByQuery(this._webUrl, this.LIST_TITLE_DRIVERS, q);
  }

  private _checkEntry() {
    this._getUserEntry().then(entryResponse => {
      let hasEntry = entryResponse.value.length > 0;
      this.setState(() => {
        return {
          userHasEntry: hasEntry,
          userEntryChecked: true,
          userEntry_P1: hasEntry ? entryResponse.value[0].FieldValuesAsText["Entry_x005f_P1"] : null,
          userEntry_P2: hasEntry ? entryResponse.value[0].FieldValuesAsText["Entry_x005f_P2"] : null,
          userEntry_P3: hasEntry ? entryResponse.value[0].FieldValuesAsText["Entry_x005f_P3"] : null,
          userEntry_P4: hasEntry ? entryResponse.value[0].FieldValuesAsText["Entry_x005f_P4"] : null,
          userEntry_P5: hasEntry ? entryResponse.value[0].FieldValuesAsText["Entry_x005f_P5"] : null,
          userEntryId: hasEntry ? entryResponse.value[0].Id : null
        };
      });
    });
  }

  @autobind
  private _driverSelected(fieldName: string, valSelected) {
    {
      this.setState((prevState) => {
        var allSelected: Array<number> = prevState.allselectedKeys;
        var selectionsValid: boolean = false;
        var validationError: string = prevState.validationError ? prevState.validationError : "";
        if (prevState[fieldName]) { //remove old value
          allSelected.splice(allSelected.indexOf(prevState[fieldName].key), 1);
        }

        if (allSelected.filter(x => x === valSelected.key).length > 0) {
          validationError = "Please select unique names. Duplicates are not allowed.";
        } else {
          validationError = "";
          if (allSelected.filter((v, i, self) => { return self.indexOf(v) === i }).length == 4) {
            selectionsValid = true;
          }
        }
        allSelected.push(valSelected.key);

        var newState = {};
        newState[fieldName] = valSelected;
        newState["selectionsValid"] = selectionsValid;
        newState["allselectedKeys"] = allSelected;
        newState["validationError"] = validationError;
        return newState;
      });
    }

  }

  @autobind
  private _showEntryForm() {
    this.setState((prevState, props) => {
      return {
        showEntryForm: !prevState.showEntryForm
      };
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
    };
    this._listService.createItem(this._webUrl, this.LIST_TITLE_ENTRIES, selectedData).then(() => {
      this._checkEntry();
      this.setState((prevState, props) => {
        return {
          showEntryForm: false
        };
      });
    });
  }

  @autobind
  private _showDeleteDialog() {
    this.setState((prevState, props) => {
      return {
        showDeleteDialog: !prevState.showDeleteDialog
      };
    }
    );

  }

  @autobind
  private _removeEntry() {
    this._listService.deleteItem(this._webUrl, this.LIST_TITLE_ENTRIES, this.state.userEntryId).then(() => {
      this._checkEntry();
      this._showDeleteDialog();
    });

  }
}
