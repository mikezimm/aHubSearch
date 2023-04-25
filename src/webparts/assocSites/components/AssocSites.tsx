import * as React from 'react';
import styles from './AssocSites.module.scss';
import { IAssocSitesProps, IAssocSitesState } from './IAssocSitesProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient, IHttpClientOptions, SPHttpClientResponse } from '@microsoft/sp-http';
// import { getAssocSites } from '@mikezimm/fps-library-v2/lib/pnpjs/Hubs/getAssocSites';
// import { getAssociatedSitesTest, IAssocHubsErrorObj } from '@mikezimm/fps-pnp2/lib/services/sp/hubs/getAssocSitesTest';


export default class AssocSites extends React.Component<IAssocSitesProps, IAssocSitesState> {

  public constructor(props:IAssocSitesProps){
    super(props);

   //  const urlVars : any = this.props.urlVars;
   //  const debugMode = urlVars.debug === 'true' ? true : false;

    this.state = {
      sites: null,
      response: null,
     };
  }

    /***
  *     .o88b.  .d88b.  .88b  d88. d8888b.      d8888b. d888888b d8888b.      .88b  d88.  .d88b.  db    db d8b   db d888888b 
  *    d8P  Y8 .8P  Y8. 88'YbdP`88 88  `8D      88  `8D   `88'   88  `8D      88'YbdP`88 .8P  Y8. 88    88 888o  88 `~~88~~' 
  *    8P      88    88 88  88  88 88oodD'      88   88    88    88   88      88  88  88 88    88 88    88 88V8o 88    88    
  *    8b      88    88 88  88  88 88~~~        88   88    88    88   88      88  88  88 88    88 88    88 88 V8o88    88    
  *    Y8b  d8 `8b  d8' 88  88  88 88           88  .8D   .88.   88  .8D      88  88  88 `8b  d8' 88b  d88 88  V888    88    
  *     `Y88P'  `Y88P'  YP  YP  YP 88           Y8888D' Y888888P Y8888D'      YP  YP  YP  `Y88P'  ~Y8888P' VP   V8P    YP    
  *                                                                                                                          
  *                                                                                                                          
  */

  public async componentDidMount(): Promise<void> {

    const options: any = {
      getNoMetadata: {
        // headers: { 'ACCEPT' : 'application/json; odata.metadata=none' }
        headers: { 'ACCEPT': 'application/json; odata=nometadata' }
      }
    }

    console.log( await this.getSearchResults( this.props.context.pageContext.web.absoluteUrl , '' ));
    const departmentId = this.props.context.pageContext.legacyPageContext.departmentId;
    console.log( 'departmentId', departmentId ); // Verified getting departmentId

    const api1:  string =  `${window.location.origin}/sites/Templates/_api/search/query?`;
    // const api1:  string =  `${window.location.origin}/sites/_api/search/query?`;
    // const apiQuery: string = `querytext=%27contentclass:STS_Site%20AND%20departmentId:{${departmentId}}%27&amp;`;
    // const apiSelect: string = `selectproperties=%27Title,SiteLogo%27&amp;`;
    // const apiOthers: string = `trimduplicates=false&amp;clienttype=%27ContentSearchRegular%27`;

    // const fullApi: string = `${ api1 }${ apiQuery }${ apiSelect }${ apiOthers }`;
    const fullApi: string = `${ api1 }querytext='sharepoint'`;
    // const fullApi: string = `${window.location.origin}/sites/Templates/_api/web/title`;

    console.log( 'apiEndPoint', fullApi ); // pasting this value into the browser will give me the results.

    const response: any = await this.props.spHttpClient.get(fullApi, SPHttpClient.configurations.v1 , options.getNoMetadata );
    const stateResponse: any =  await response.json();
    
    console.log( 'response', stateResponse );

    this.setState({ response: stateResponse });

 }

 	/**************************************************************************************************
	 * Recursively executes the specified search query using batches of 500 results until all results are fetched
	 * @param webUrl : The web url from which to call the search API
	 * @param queryParameters : The search query parameters following the "/_api/search/query?" part
	 * @param startRow : The row from which the search needs to return the results from
	 **************************************************************************************************/
   public getSearchResults(webUrl: string, queryParameters: string, ): Promise<any> {
		return new Promise<any>((resolve,reject) => {

      const api1:  string =  `${webUrl}/_api/search/query?querytext=`;
      const fullApi: string = `${ api1 }'sharepoint'`;

			this.props.spHttpClient.get(fullApi, SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
				if(response.ok) {
					resolve(response.json());
				}
				else {
					reject(response.statusText);
				}
			})
			.catch((error) => { reject(error); }); 
		});	
	}

  public render(): React.ReactElement<IAssocSitesProps> {
    const {
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;
    const departmentId = this.props.context.pageContext.legacyPageContext.departmentId;
    return (
      <section className={`${styles.assocSites} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
        </div>
        <div>
          <h3>Your Hubsite Id is { departmentId }</h3>
          <div>{ JSON.stringify( this.state.response ) }</div>
        </div>
      </section>
    );
  }
}
