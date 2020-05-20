import * as React from "react";
import { Button, ButtonType } from "office-ui-fabric-react";
import Header from "./Header";
import CustomerDetailsList from "./CustomerDetailslist"
import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";
import { GraphClient } from "../../dal/GraphClient";
import { IEmail } from "../../model/dto/IEmail";
import { SPClient } from "../../dal/SPClient";
import { ICustomerRecord } from "../../model/dto/ICustomerRecord";
import { Announced } from 'office-ui-fabric-react/lib/Announced';
// import { TextField, ITextFieldStyles } from 'office-ui-fabric-react/lib/TextField';
import { DetailsList, DetailsListLayoutMode, Selection, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { MarqueeSelection } from 'office-ui-fabric-react/lib/MarqueeSelection';
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
import { mergeStyles } from 'office-ui-fabric-react/lib/Styling';
/* global Button, Header, HeroList, HeroListItem, Progress */

const exampleChildClass = mergeStyles({
  display: 'block',
  marginBottom: '10px',
});

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  listItems: HeroListItem[];
  customerRelatedProducts: ICustomerRecord[]
  userDisplayName?: string;
  relatedEmails?: IEmail[];
  relatedSites?: any[];
  customerDetailsLoading:boolean;
}
export default class App extends React.Component<AppProps, AppState> {
  protected graphClient: GraphClient = new GraphClient();
  protected spClient: SPClient = new SPClient();
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [],
      customerRelatedProducts:[],
      customerDetailsLoading:false
    };
    OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true, forMSGraphAccess: true });
  }

  async componentDidMount() {
    let userProfile = await this.graphClient.getMyProfileInformation();
    let userDisplayName = userProfile.displayName;
    this.setState({
      listItems: [
        {
          icon: "Ribbon",
          primaryText: "Achieve more with Office integration"
        },
        {
          icon: "Unlock",
          primaryText: "Unlock features and functionality"
        },
        {
          icon: "Design",
          primaryText: "Create and visualize like a pro"
        }
      ],
      userDisplayName
    });
  }
  
  click = async () => {
    let subject = Office.context.mailbox.item.normalizedSubject;
    let emails = await this.graphClient.searchMyMailbox(subject);
    this.setState({
      relatedEmails: emails
    });
  };
  searchSP = async ()=>{
    let fromEmail = Office.context.mailbox.item.from;
    let customerRecords = await this.spClient.getProductsRelatedCustomer(fromEmail.emailAddress);
    this.setState({
      customerRelatedProducts : customerRecords
    })
  }
  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress title={title} logo="assets/logo-filled.png" message="Please sideload your addin to see app body." />
      );
    }

    return (
      <div className="ms-welcome">
        <Header logo="assets/logo-filled.png" title={this.props.title} message={"Welcome " + this.state.userDisplayName} />
          <Button
            className="ms-welcome__action"
            buttonType={ButtonType.hero}
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.click.bind(this)}
          >
            Search similar emails
          </Button>
        {this.state.relatedEmails && this.state.relatedEmails.map(email =>{
          return <div>
            Subject: {email.subject}
            Sender: {email.sender.emailAddress.name}
            Preview: {email.bodyPreview}
          </div>;
        })}
       
      <CustomerDetailsList spClient={this.spClient}></CustomerDetailsList>
      </div>
    );
  }
}
