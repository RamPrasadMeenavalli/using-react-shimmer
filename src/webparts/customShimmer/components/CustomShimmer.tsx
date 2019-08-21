import * as React from 'react';
import styles from './CustomShimmer.module.scss';
import { ICustomShimmerProps } from './ICustomShimmerProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { Shimmer, ShimmerElementsGroup, ShimmerElementType } from 'office-ui-fabric-react/lib/Shimmer';
import {sp, SearchQuery, SearchResults} from '@pnp/sp';

export interface ICustomShimmerState{
  loaded: boolean;
  sites: any[];
}

export default class CustomShimmer extends React.Component<ICustomShimmerProps, ICustomShimmerState> {
  public constructor(props:ICustomShimmerProps, state:ICustomShimmerState){
    super(props);
    // Set the intial state with 5 dummy values for sites property
    this.state = {
      loaded: false,
      sites:[
        {Title:"Site 1"},
        {Title:"Site 2"},
        {Title:"Site 3"},
        {Title:"Site 4"},
        {Title:"Site 5"},
      ]
    };
  }

  public render(): React.ReactElement<ICustomShimmerProps> {
    // Create the structure for each site element wrapped within a Shimmer component
    const elements = this.state.sites.map((val,index) => {
      return <div style={{padding:"10px", background:"white"}}>
      <Shimmer
        customElementsGroup={this._getElementsForSiteListing()}
        isDataLoaded={this.state.loaded}>
            <div className={ styles.siteRow }>
            <div className={ styles.imgColumn }>
              <img src={val.SiteLogo} height={40} width={40}/>
            </div>
            <div className={ styles.titleColumn }>
              {val.Title}
            </div>
          </div>
      </Shimmer>
      </div>;
    });
    
    return (<div className={ styles.customShimmer }>
      <div className={ styles.container }>
        <div className={ styles.row }>
          <div className={ styles.column }>

            {elements}

          </div>
        </div>
      </div>
    </div>);
  }

  // This method provides the structure of the Shimmer with elements and groups
  private _getElementsForSiteListing= (): JSX.Element => {
    return (
      <div
        style={{ display: 'flex' }}
      >
        <ShimmerElementsGroup
          shimmerElements={[
            { type: ShimmerElementType.line, width: 40, height: 40 },
            { type: ShimmerElementType.gap, width: 10, height: 40 }
          ]}
        />
        <ShimmerElementsGroup
          flexWrap={true}
          shimmerElements={[
            { type: ShimmerElementType.gap, width: 370, height: 10 },
            { type: ShimmerElementType.line, width: 370, height: 10 },
            { type: ShimmerElementType.gap, width: 370, height: 10 }
          ]}
        />
      </div>
    );
  }

  // Query the search API to get the list of sites and set them in the state
  public componentDidMount()
  {
    // Sometime PnPjs is very fast, that we cannot see the shimmer.
    // So let's put a 5 seconds delay to see the Shimmer effect
    setTimeout(() => { sp.search({
      SelectProperties: ["Title","SiteLogo"],
      Querytext: `contentclass:STS_Site AND WebTemplate:'Group'`,
      RowLimit: 5
    }).then(w => {
      console.dir(w);
      this.setState({
        sites:w.PrimarySearchResults,
        loaded:true
      });
    });}, 5000);
    
  }
}
