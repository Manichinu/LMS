import * as React from 'react';
import styles from './Aboutus.module.scss';
import { IAboutusProps } from './IAboutusProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPComponentLoader } from "@microsoft/sp-loader";
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import './reactAccordion.css';
import { Web } from '@pnp/sp/webs';
import {
  Accordion,
  AccordionItem,
  AccordionItemHeading,
  AccordionItemButton,
  AccordionItemPanel,
} from 'react-accessible-accordion';

export interface IAboutusState {
  items: Array<any>;
  allowMultipleExpanded: boolean;
  allowZeroExpanded: boolean;

}
let NewWeb:any;
 NewWeb = Web("https://tmxin.sharepoint.com/sites/POC/SPIP/");
export default class Aboutus extends React.Component<IAboutusProps,IAboutusState> {

  public constructor(props: IAboutusProps) {
    super(props);

    this.state = {
      items: new Array<any>(),
      allowMultipleExpanded: this.props.allowMultipleExpanded,
      allowZeroExpanded: this.props.allowZeroExpanded
    };
    this.getListItems();
  
  SPComponentLoader.loadCss(
    `https://maxcdn.bootstrapcdn.com/bootstrap/3.4.1/css/bootstrap.min.css`
  );
  SPComponentLoader.loadCss(`https://fonts.googleapis.com`);
  SPComponentLoader.loadCss(`https://fonts.gstatic.com" crossorigin`);
  SPComponentLoader.loadCss(`https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700&display=swap`);
 
  SPComponentLoader.loadCss(
    `https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css`
  );

  SPComponentLoader.loadScript(
    `https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js`
  );

  SPComponentLoader.loadScript(
    `https://maxcdn.bootstrapcdn.com/bootstrap/3.4.1/css/bootstrap.min.js`
  );
  SPComponentLoader.loadCss(
    `https://maxcdn.bootstrapcdn.com/bootstrap/3.4.1/css/bootstrap.min.css`
  );
  SPComponentLoader.loadCss(
    `https://tmxin.sharepoint.com/sites/POC/SPIP/SiteAssets/LeavePortal/css/style.css?v=1.6`
  );
  SPComponentLoader.loadCss('//code.jquery.com/ui/1.11.4/themes/smoothness/jquery-ui.css')
  }

  public componentDidUpdate(prevProps:IAboutusProps): void {
    if(prevProps.listId !== this.props.listId) {
      this.getListItems();
    }

    if(prevProps.allowMultipleExpanded !== this.props.allowMultipleExpanded || prevProps.allowZeroExpanded !== this.props.allowZeroExpanded) {
      this.setState({
        allowMultipleExpanded: this.props.allowMultipleExpanded,
        allowZeroExpanded: this.props.allowZeroExpanded
      });
    }
  }
  private getListItems(): void {
    if(typeof this.props.listId !== "undefined" && this.props.listId.length > 0) {
      NewWeb.lists.getByTitle("LeaveTypeCollection").items.select("Types","Details").get()
        .then((results: Array<any>) => {
          this.setState({
            items: results
          });
        })
        .catch((error:any) => {
          console.log("Failed to get list items!");
          console.log(error);
        });
    }
  }
  public render(): React.ReactElement<IAboutusProps> {


    const listSelected:boolean = typeof this.props.listId !== "undefined" && this.props.listId.length > 0;
    const { allowMultipleExpanded, allowZeroExpanded } = this.state;
    return (
      <div className={ styles.aboutus }>
         <header>
    <div className="container">
        <div className="logo">
            <img src="https://tmxin.sharepoint.com/sites/POC/SPIP/SiteAssets/LeavePortal/img/logo_small.png" alt="image" />
            <ul>
               {/* <li> <a href=""> <img src="https://tmxin.sharepoint.com/sites/POC/SPIP/SiteAssets/LeavePortal/img/settings.svg" alt="image"/> </a> </li>
                <li> <a href="" className="relative"> <img src="https://tmxin.sharepoint.com/sites/POC/SPIP/SiteAssets/LeavePortal/img/notification.svg" alt="image"/>  <span className="noto-count">  2  </span> </a> </li>
    <li className="person-details">  <img src="https://tmxin.sharepoint.com/sites/POC/SPIP/SiteAssets/LeavePortal/img/user.png" alt="image"/> <span> Mohammed </span> <img src="https://tmxin.sharepoint.com/sites/POC/SPIP/SiteAssets/LeavePortal/img/down.svg" alt="image"/>  </li>*/}
            </ul>
        </div>
    </div>
 </header>
        {!listSelected &&
          <Placeholder
            iconName='MusicInCollectionFill'
            iconText='Configure your web part'
            description='Select a list with a Title field and Content field to have its items rendered in a collapsible accordion format'
            buttonLabel='Choose a List'
            onConfigure={this.props.onConfigure} />
        }
      {listSelected &&
        <div>
          <WebPartTitle displayMode={this.props.displayMode}
                title={this.props.accordionTitle}
                updateProperty={this.props.updateProperty}
                 />
          <Accordion allowZeroExpanded={allowZeroExpanded} allowMultipleExpanded={allowMultipleExpanded}>
            {this.state.items.map((item:any) => {
              return (
                <AccordionItem>
                  <AccordionItemHeading>
                    <AccordionItemButton>
                      {item.Types}
                    </AccordionItemButton>
                  </AccordionItemHeading>
                    <AccordionItemPanel>
                      <p  dangerouslySetInnerHTML={{__html: item.Details}} />
                    </AccordionItemPanel>
                </AccordionItem>
                );
              })
            }
          </Accordion>
        </div>
  }
      </div>
          
   
      );
          
  }
}
   
