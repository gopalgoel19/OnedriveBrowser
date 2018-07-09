import './App.css';
import * as React from 'react';
import * as Moment from 'react-moment';
import * as axios from 'axios';
import {
  DocumentCard,
  DocumentCardActivity,
  DocumentCardPreview,
  DocumentCardTitle
} from 'office-ui-fabric-react/lib/DocumentCard';
import { initializeIcons } from '@uifabric/icons';

// Register icons and pull the fonts from the default SharePoint cdn:
initializeIcons();

import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import {
  DetailsList,
  DetailsListLayoutMode,
  Selection,
  SelectionMode
} from 'office-ui-fabric-react/lib/DetailsList';
import { Link } from 'office-ui-fabric-react/lib/Link';


import { authContext } from './adalConfig';
import { adalApiFetch } from './adalConfig';



const columns = [
  { 
    key: "icon", 
    name: "...", 
    fieldName: "icon", 
    minWidth: 20, 
    maxWidth: 20,
    isResizable: true,  
    // onRender: (item) => (
    //   <img src={item.icon}/>
    // )
    onRender: (item) => {
      if(item.icon == ''){
        return <i className="ms-Icon ms-Icon--FabricFolderFill" aria-hidden="true"></i>
      }
      else{
        return <img src={item.icon}/>
      }
    }
  },
  { 
    key: "name", 
    name: "Name", 
    fieldName: "name", 
    minWidth: 20, 
    maxWidth: 300,
    isResizable: true,
    onRender: (item) => (
      <Link href={item.url} target='_blank' className="ms-font-m" style={{textDecoration:'none', color: 'black'}}>{item.name}</Link>
    )
  },
  { 
    key: "position", 
    name: "Modified",
    fieldName: 'position',
    minWidth: 20, 
    maxWidth: 300,
    isResizable: true
  },
  { 
    key: "office", 
    name: "Modified By",
    fieldName: 'office',
    minWidth: 160, 
    maxWidth: 280,
    isResizable: true,
    // onRender: (item: any) => (
    //   <div className='vstack'>
    //     <PrimaryButton
    //       text={ `Going to ${item.office}` }
    //     />
    //     <DefaultButton
    //       text={ `In ${item.office}` }
    //     />
    //     <DefaultButton
    //       text={ `Leaving ${item.office}` }
    //     />
        
    //   </div>
    // )         
  },
  {
    key: "another",
    name: 'Size',
    fieldName: 'another',
    minWidth: 70,
    maxWidth: 100,
    // onRender: (item) => (
    //   <Link href='//www.microsoft.com'>I am a link</Link>
    // )
  }
];
             
const fileIcons: { name: string }[] = [
  { name: 'accdb' },
  { name: 'csv' },
  { name: 'docx' },
  { name: 'dotx' },
  { name: 'mpp' },
  { name: 'mpt' },
  { name: 'odp' },
  { name: 'ods' },
  { name: 'odt' },
  { name: 'one' },
  { name: 'onepkg' },
  { name: 'onetoc' },
  { name: 'potx' },
  { name: 'ppsx' },
  { name: 'pptx' },
  { name: 'pub' },
  { name: 'vsdx' },
  { name: 'vssx' },
  { name: 'vstx' },
  { name: 'xls' },
  { name: 'xlsx' },
  { name: 'xltx' },
  { name: 'xsn' }
];

// console.log(items);

class App extends React.Component<{},{items: Array<any>}> {
  constructor(props){
    super(props);
    
    this.state = {
      items: []
    };


    adalApiFetch(fetch, 'https://graph.microsoft.com/v1.0/me/drive/root/children', {}).then((response) => {
        // console.log(response);
        // This is where you deal with your API response. In this case, we            
        // interpret the response as JSON, and then call `setState` with the
        // pretty-printed JSON-stringified object.
        response.json().then((response) => {
            // console.log(responseJson.body);
            console.log(JSON.stringify(response, null, 2));
            for (let i = 0; i < response.value.length; i++) { 
              let value = response.value[i];
              let size: any = value.size/1024;
              let d: any = new Date(value.lastModifiedDateTime);
              d = d.toString().slice(0,25);
              // console.log(d);
              if(size>=1024){
                size = size/1024;
                size = size.toFixed(2) + " MB";
              }
              else{
                size = size.toFixed(2) + " KB"
              }
              let item = {
                icon: '',
                key: 'item' + i,
                name: value.name,
                position: d,
                office: value.lastModifiedBy.user.displayName,
                another: size,
                index: i,
                url: ''
              };

              let icon = '';
              let name = value.name;
              for (let i=0; i< fileIcons.length ; i++){
                if(name.endsWith('.' + fileIcons[i].name)){
                  icon = `https://static2.sharepointonline.com/files/fabric/assets/brand-icons/document/svg/${fileIcons[i].name}_16x1.svg`;
                }
              }
              if(icon == ''){
                  // icon = `https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/svg/onedrive_16x1.svg`;
                  
              }
              item.icon = icon;
              item.url = value.webUrl;
              this.setState((prevState) => ({
                items: prevState.items.concat(item)
              }));
              // this.state.items.push();
            }
        });
      }).catch((error) => {
        // Don't forget to handle errors!
        console.error(error);
      });

    // this.clickHandler = new Selection({
    //   onSelectionChanged: () => {
    //     this.setState({
    //       selectionDetails: this._getSelectionDetails(),
    //       isModalSelection: this._selection.isModal()
    //     });
    //   }
    // });
  
  }

  public render() {
    return (
        <div>
          <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.0.0/css/fabric.min.css"/>
          <p className="ms-font-su">OneDrive</p>
          <DetailsList 
            items={ this.state.items }
            columns={ columns }
            selectionMode= {SelectionMode.none}
          />
        </div>
    );
  }
}






export default App;





// ReactDOM.render(<MyPage />, document.body.firstChild);