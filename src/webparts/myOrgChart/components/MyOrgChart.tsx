import * as React from 'react';
import styles from './MyOrgChart.module.scss';
import { IMyOrgChartProps, IMyOrgChartState, IPersons, IPerson } from '.';
import { WebPartTitle } from '@pnp/spfx-controls-react/lib/WebPartTitle';
import * as strings from 'MyOrgChartWebPartStrings';
import { Persona, PersonaSize, PersonaPresence } from 'office-ui-fabric-react/lib/Persona';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/components/Spinner';
import { List } from 'office-ui-fabric-react/lib/components/List';
import { Person } from './person';

interface IPersonaListProps {
  title: string;
  users: IPerson[];
  //getProfilePhoto: (photoUrl: string) => string;
  //onProfileLinkClick: (profileLink: string) => void;
}

class PersonaList extends React.Component<IPersonaListProps, {}> {
  public render() {
    return (
      <div>
        <div className={styles.subTitle}>{this.props.title}</div>
        {this.props.users.map((user, index) => (
          <div key={index}>
            <Persona
              //imageUrl={this.props.getProfilePhoto(user.PictureUrl)}
              primaryText={user.displayName}
              secondaryText={user.jobTitle}
              size={PersonaSize.regular}
              presence={PersonaPresence.none}
              //onClick={() => this.props.onProfileLinkClick(user.UserUrl)}
            />
          </div>
        ))}
      </div>
    );
  }
}

export class MyOrgChart extends React.Component<IMyOrgChartProps, IMyOrgChartState> {
  
  constructor(props: IMyOrgChartProps, state: IMyOrgChartState) {
    super(props);

    this.state = {
      manager: null,
      loadingMgr: true,
      errorMgr: null,
      user: null,
      loadingUser: true,
      errorUser: null,
      reports: [],
      loadingReps: true,
      errorReps: null
    };
  }

  /**
   * Fetch the current user's manager and direct reports
   */
  private _fetchUserOrgChart(): void {
    console.log("starting fetch user org chart");
    if (this.props.graphClient) {
      this.setState({
        errorMgr: null,
        errorReps: null,
        errorUser: null,
        loadingMgr: true,
        loadingReps: true,
        loadingUser: true,
      });

      console.log("fetching user mgr");
      this.props.graphClient
        .api("me/manager")
        .version("v1.0")
        .select("id,displayName,jobTitle,businessPhones,mail")
        .get((err, res: IPerson) => {
          if (err) {
            // Something failed calling the MS Graph
            // known issue for me/manager is that an error will be returned if the user does not have a manager
            if (err.code != 'Request_ResourceNotFound') {
              this.setState({
                errorMgr: err.message ? err.message : strings.Error,
                manager: null,
                loadingMgr: false
              });
            }
          }

          // Check if a response was retrieved
          if (res) {
            this.setState({
              manager: {
                "@odata.type": res["@odata.type"],
                id: res.id,
                displayName: res.displayName,
                jobTitle: res.jobTitle,
                mail: res.mail,
                businessPhones: res.businessPhones
              },
              loadingMgr: false
            });
          } else {
            // No manager retrieved
            this.setState({
              manager: null,
              loadingMgr: false
            });
          }
        });

      console.log("fetching user data");
      this.props.graphClient
        .api("me")
        .version("v1.0")
        .select("id,displayName,jobTitle,businessPhones,mail")
        .get((err, res: IPerson) => {
          if (err) {
            // Something failed calling the MS Graph
            this.setState({
              errorUser: err.message ? err.message : strings.Error,
              user: null,
              loadingUser: false
            });
          }

          if (res) {
            this.setState({
              user: {
                "@odata.type": res["@odata.type"],
                id: res.id,
                displayName: res.displayName,
                jobTitle: res.jobTitle,
                mail: res.mail,
                businessPhones: res.businessPhones
              },
              loadingUser: false
            });
          }
        });

      // get direct reports
      console.log("fetching user's direct reports");
      this.props.graphClient
        .api("me/directReports")
        .version("v1.0")
        .select("id,displayName,jobTitle,businessPhones,mail")
        .get((err, res: IPersons) => {
          if (err) {
            // Something failed calling the MS Graph
            this.setState({
              errorReps: err.message ? err.message : strings.Error,
              reports: [],
              loadingReps: false
            });
          }

          if (res && res.value && res.value.length > 0) {
            this._processPeoplesResults(res.value);
          } else {
            // No documents retrieved
            this.setState({
              reports: [],
              loadingReps: false
            });
          }
        });
    }
  }

  /**
   * Process the retrieved MS Graph Contacts
   * @param contacts
   */
  private _processPeoplesResults(persons: IPerson[]): void {
    this.setState({
      reports: persons,
      loadingReps: false
    });
  }

  /**
   * Renders the list cell for the persona's
   */
  private _onRenderCell = (item: IPerson, index: number | undefined): JSX.Element => {
    return <Person className={styles.persona} person={item} graphClient={this.props.graphClient} />;
  }

  /**
   * componentDidMount lifecycle hook
   */
  public componentDidMount(): void {
    this._fetchUserOrgChart();
  }
  
  public render(): React.ReactElement<IMyOrgChartProps> {
    return (
      <div className={ styles.myOrgChart }>
        <WebPartTitle displayMode={this.props.displayMode}
          title={this.props.title}
          updateProperty={this.props.updateProperty} />

        {
          this.state.loadingMgr && this.state.loadingReps && this.state.loadingUser ? (
           <Spinner label={strings.Loading} size={SpinnerSize.large} />
          ) : null
        }

        {
          this.state.manager ? (
            <div>
              <div>Manager</div>
              <Person className={styles.persona} person={this.state.manager} graphClient={this.props.graphClient} />
            </div>
            ) : (
              !this.state.loadingMgr && (
                this.state.errorMgr ?
                  <span className={styles.error}>{this.state.errorMgr}</span> :
                  <span className={styles.noContacts}>{strings.NoManager}</span>
              )
            )
        }
        {
          this.state.user ? (
            <div>
              <div>You</div>
              <Person className={styles.persona} person={this.state.user} graphClient={this.props.graphClient} />
            </div>
          ) : (
            <div>No user info</div>
          )
        }
        {
          this.state.reports && this.state.reports.length > 0 ? (
            <List items={this.state.reports}
              renderedWindowsAhead={4}
              onRenderCell={this._onRenderCell} />
              /*
            <div>
              <PersonaList
                title="Reports"
                users={this.state.reports}
                //getProfilePhoto={this.getProfilePhoto.bind(this)}
                //onProfileLinkClick={this.onProfileLinkClick.bind(this)}
              />
            </div> */
          ) : (
            <div>No direct reports</div>
          )
        }

      </div>
    );
  }
}
