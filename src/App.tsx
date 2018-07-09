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


// import * as React from 'react';
// import * as ReactDOM from 'react-dom';
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

// console.log(authContext);

const dateToFormat = '1976-04-19T12:59-0500';

const columns = [
  { 
    key: "name", 
    name: "Name", 
    fieldName: "name", 
    minWidth: 20, 
    maxWidth: 300,
    isResizable: true  
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
             


// console.log(items);

class App extends React.Component<{},{items: Array<any>}> {
  constructor(props){
    super(props);
    
    this.state = {
      items: []
    };


    adalApiFetch(fetch, 'https://graph.microsoft.com/v1.0/me/drive/root/children', {})
      .then((response) => {
        // console.log(response);
        // This is where you deal with your API response. In this case, we            
        // interpret the response as JSON, and then call `setState` with the
        // pretty-printed JSON-stringified object.
        response.json()
          .then((response) => {
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
                key: 'item' + i,
                name: value.name,
                position: d,
                office: value.lastModifiedBy.user.displayName,
                another: size,
                index: i
              };
              this.setState((prevState) => ({
                items: prevState.items.concat(item)
              }));
              // this.state.items.push();
            }
          });
      })
      .catch((error) => {

        // Don't forget to handle errors!
        console.error(error);
      });
      
  }
  public render() {
    return (
        <div>
          <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.0.0/css/fabric.min.css"/>
          <h1 className="ComponentPage-subHeading">Files</h1>
          <DetailsList items={ this.state.items } columns={ columns } selectionMode= {SelectionMode.none}/>
        </div>
    );
  }
}






export default App;





// ReactDOM.render(<MyPage />, document.body.firstChild);