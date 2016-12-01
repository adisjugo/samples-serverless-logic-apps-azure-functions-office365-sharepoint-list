import * as React from 'react';
import { css } from 'office-ui-fabric-react';

import styles from '../ProvisionLists.module.scss';
import { IProvisionListsWebPartProps } from '../IProvisionListsWebPartProps';
import { IWebPartContext } from '@microsoft/sp-webpart-base';

export interface IProvisionListsProps extends IProvisionListsWebPartProps {
  context : IWebPartContext
}

export interface IProvisionListsState {
  sourceListTitle : string;
  destinationListTitle : string;
}

export default class ProvisionLists extends React.Component<IProvisionListsProps, IProvisionListsState> {
  constructor(props) {
    super(props);

    this.state = { 
      sourceListTitle: 'Source List',
      destinationListTitle: 'Destination List'
    };
  }

  protected handleSourceListTitleChange = (event) : void => this.setState({ sourceListTitle : event.target.value } as IProvisionListsState);
  protected handleDestinationListTitleChange = (event) : void => this.setState({ destinationListTitle : event.target.value } as IProvisionListsState);

  protected handleProvisionSourceListClick = (event) : void => {
    let url = `${this.props.provisionSourceListEndpointUrl}&siteUrl=${this.props.context.pageContext.web.absoluteUrl}&listTitle=${this.state.sourceListTitle}`;
    this.props.context.basicHttpClient.get(url)
      .then((r) => {
        console.log(r);
      });
  }

  protected handleProvisionDestinationListClick = (event) : void => {
    let url = `${this.props.provisionDestinationListEndpointUrl}&siteUrl=${this.props.context.pageContext.web.absoluteUrl}&listTitle=${this.state.destinationListTitle}`;
    this.props.context.basicHttpClient.get(url)
      .then((r) => {
        console.log(r);
      });
  }

  public render(): JSX.Element {
    return (
      <div className={styles.provisionLists}>
        <div className={styles.container}>
          <div className={css('ms-Grid-row ms-bgColor-themeDark ms-fontColor-white', styles.row)}>
            <div className='ms-Grid-col ms-u-lg12'>
              <span className='ms-font-xl ms-fontColor-white'>
                Provision Lists
              </span>
              <div className='section'>
                <div>
                  <label for='sourceListTitle'>Source List Title: </label>
                </div>
                <div>
                  <input type='text' id='sourceListTitle' value={this.state.sourceListTitle} onChange={this.handleSourceListTitleChange} />
                </div>
                <div>
                  <a className={css('ms-Button', styles.button)} onClick={this.handleProvisionSourceListClick}>
                    <span className='ms-Button-label'>Provision</span>
                  </a>
                </div>
              </div>
              <div className='section'>
                <div>
                  <label for='destinationListTitle'>Destination List Title: </label>
                </div>
                <div>
                  <input type='text' id='destinationListTitle' value={this.state.destinationListTitle} onChange={this.handleDestinationListTitleChange} />
                </div>
                <div>
                  <a className={css('ms-Button', styles.button)} onClick={this.handleProvisionDestinationListClick}>
                    <span className='ms-Button-label'>Provision</span>
                  </a>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
