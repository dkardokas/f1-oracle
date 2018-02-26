import * as React from 'react';
import styles from './F1Entry.module.scss';
import { IF1EntryProps } from './IF1EntryProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class F1Entry extends React.Component<IF1EntryProps, {}> {
  public render(): React.ReactElement<IF1EntryProps> {
    return (
      <div className={ styles.f1Entry }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
              <p className={ styles.description }>{escape(this.props.description)}</p>
              <a href="https://aka.ms/spfx" className={ styles.button }>
                <span className={ styles.label }>Learn more</span>
              </a>
            </div>
          </div>
          <div id="spListContainer" />
        </div>
      </div>
    );
  }
}
