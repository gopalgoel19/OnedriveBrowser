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
    name: "Last Modified",
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

// const _columns= [
//   {
//     key: 'column1',
//     name: 'File Type',
//     headerClassName: 'DetailsListExample-header--FileIcon',
//     className: 'DetailsListExample-cell--FileIcon',
//     iconClassName: 'DetailsListExample-Header-FileTypeIcon',
//     iconName: 'Page',
//     isIconOnly: true,
//     fieldName: 'name',
//     minWidth: 16,
//     maxWidth: 16,
//     onColumnClick: this._onColumnClick,
//     onRender: (item) => {
//       return <img src={item.iconName} className={'DetailsListExample-documentIconImage'} />;
//     }
//   },
//   {
//     key: 'column2',
//     name: 'Name',
//     fieldName: 'name',
//     minWidth: 210,
//     maxWidth: 350,
//     isRowHeader: true,
//     isResizable: true,
//     isSorted: true,
//     isSortedDescending: false,
//     onColumnClick: this._onColumnClick,
//     data: 'string',
//     isPadded: true
//   },
//   {
//     key: 'column3',
//     name: 'Date Modified',
//     fieldName: 'dateModifiedValue',
//     minWidth: 70,
//     maxWidth: 90,
//     isResizable: true,
//     onColumnClick: this._onColumnClick,
//     data: 'number',
//     onRender: (item) => {
//       return <span>{item.dateModified}</span>;
//     },
//     isPadded: true
//   },
//   {
//     key: 'column4',
//     name: 'Modified By',
//     fieldName: 'modifiedBy',
//     minWidth: 70,
//     maxWidth: 90,
//     isResizable: true,
//     isCollapsable: true,
//     data: 'string',
//     onColumnClick: this._onColumnClick,
//     onRender: (item) => {
//       return <span>{item.modifiedBy}</span>;
//     },
//     isPadded: true
//   },
//   {
//     key: 'column5',
//     name: 'File Size',
//     fieldName: 'fileSizeRaw',
//     minWidth: 70,
//     maxWidth: 90,
//     isResizable: true,
//     isCollapsable: true,
//     data: 'number',
//     onColumnClick: this._onColumnClick,
//     onRender: (item) => {
//       return <span>{item.fileSize}</span>;
//     }
//   }
// ];


let positions = [ 'Gopal Goel', 'Heramb Mathkar', 'Paras Jindal', 'Shashank Shekar', 'Shivang Bansal'];

let offices = [ 'Seattle', 'New York', 'Tokyo', 'California'];

function randWord(words) {
  return words[Math.floor(Math.random() * words.length)];
}

let items : any = [];

const url = 'https://graph.microsoft.com/v1.0/me/drive/root/children';

fetch(url,{
  method: 'GET',
  headers: {
    Authorization: 'Bearer eyJ0eXAiOiJKV1QiLCJub25jZSI6IkFRQUJBQUFBQUFEWDhHQ2k2SnM2U0s4MlRzRDJQYjdyYnIzWVZhTi1YSUloeEMwWUJGSkc5UDhIT2NWdmNGc1dJSHNCa0Vrcmt3TloyRlpvQkpvaG9VLUcyNll4ZzFIRV9jeDVaSDRDUnhGUVhwOEo4ai1HclNBQSIsImFsZyI6IlJTMjU2IiwieDV0IjoiVGlvR3l3d2xodmRGYlhaODEzV3BQYXk5QWxVIiwia2lkIjoiVGlvR3l3d2xodmRGYlhaODEzV3BQYXk5QWxVIn0.eyJhdWQiOiJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20iLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC83MmY5ODhiZi04NmYxLTQxYWYtOTFhYi0yZDdjZDAxMWRiNDcvIiwiaWF0IjoxNTMwODU5MDQ4LCJuYmYiOjE1MzA4NTkwNDgsImV4cCI6MTUzMDg2Mjk0OCwiYWNyIjoiMSIsImFpbyI6IkFVUUF1LzhIQUFBQXM1bDRPcFlIODFXSHlucm0xZVJJWEoxSThyMVFYck9SMUVXVEx1ZG5uR1h6SjhmczhLZ1d3TGdMK0k0eEtQSU9JLzZWREZ4eTVSbnJnVER3ZzVpYUN3PT0iLCJhbXIiOlsicnNhIiwibWZhIl0sImFwcF9kaXNwbGF5bmFtZSI6IkdyYXBoIGV4cGxvcmVyIiwiYXBwaWQiOiJkZThiYzhiNS1kOWY5LTQ4YjEtYThhZC1iNzQ4ZGE3MjUwNjQiLCJhcHBpZGFjciI6IjAiLCJkZXZpY2VpZCI6IjNhMTE4MGNlLTlmMWUtNGUxNi1iNzY3LTM4NGQ3NDY0Y2ZlMSIsImZhbWlseV9uYW1lIjoiR29lbCIsImdpdmVuX25hbWUiOiJHb3BhbCIsImlwYWRkciI6IjE2Ny4yMjAuMjM4LjE1OSIsIm5hbWUiOiJHb3BhbCBHb2VsIiwib2lkIjoiY2U0MGFkZDMtN2VkMy00NTkwLWE1OTMtNjljYWNhZGViOTlkIiwib25wcmVtX3NpZCI6IlMtMS01LTIxLTIxNDY3NzMwODUtOTAzMzYzMjg1LTcxOTM0NDcwNy0yMzY4Mzg0IiwicGxhdGYiOiIzIiwicHVpZCI6IjEwMDMwMDAwQUI0RjcyRTYiLCJzY3AiOiJDYWxlbmRhcnMuUmVhZFdyaXRlIENvbnRhY3RzLlJlYWRXcml0ZSBEZXZpY2VNYW5hZ2VtZW50QXBwcy5SZWFkLkFsbCBEZXZpY2VNYW5hZ2VtZW50QXBwcy5SZWFkV3JpdGUuQWxsIERldmljZU1hbmFnZW1lbnRDb25maWd1cmF0aW9uLlJlYWQuQWxsIERldmljZU1hbmFnZW1lbnRDb25maWd1cmF0aW9uLlJlYWRXcml0ZS5BbGwgRGV2aWNlTWFuYWdlbWVudE1hbmFnZWREZXZpY2VzLlByaXZpbGVnZWRPcGVyYXRpb25zLkFsbCBEZXZpY2VNYW5hZ2VtZW50TWFuYWdlZERldmljZXMuUmVhZC5BbGwgRGV2aWNlTWFuYWdlbWVudE1hbmFnZWREZXZpY2VzLlJlYWRXcml0ZS5BbGwgRGV2aWNlTWFuYWdlbWVudFJCQUMuUmVhZC5BbGwgRGV2aWNlTWFuYWdlbWVudFJCQUMuUmVhZFdyaXRlLkFsbCBEZXZpY2VNYW5hZ2VtZW50U2VydmljZUNvbmZpZy5SZWFkLkFsbCBEZXZpY2VNYW5hZ2VtZW50U2VydmljZUNvbmZpZy5SZWFkV3JpdGUuQWxsIERpcmVjdG9yeS5BY2Nlc3NBc1VzZXIuQWxsIERpcmVjdG9yeS5SZWFkV3JpdGUuQWxsIEZpbGVzLlJlYWRXcml0ZS5BbGwgR3JvdXAuUmVhZFdyaXRlLkFsbCBJZGVudGl0eVJpc2tFdmVudC5SZWFkLkFsbCBNYWlsLlJlYWRXcml0ZSBNYWlsYm94U2V0dGluZ3MuUmVhZFdyaXRlIE5vdGVzLlJlYWRXcml0ZS5BbGwgUGVvcGxlLlJlYWQgUmVwb3J0cy5SZWFkLkFsbCBTaXRlcy5SZWFkV3JpdGUuQWxsIFRhc2tzLlJlYWRXcml0ZSBVc2VyLlJlYWRCYXNpYy5BbGwgVXNlci5SZWFkV3JpdGUgVXNlci5SZWFkV3JpdGUuQWxsIiwic2lnbmluX3N0YXRlIjpbImR2Y19tbmdkIiwiZHZjX2RtamQiLCJrbXNpIl0sInN1YiI6InRtLWFnRGU0WTZiam1nZGNqcWw3bE5HdHFnUjVDRGxCLWE5V2lpRjNESTgiLCJ0aWQiOiI3MmY5ODhiZi04NmYxLTQxYWYtOTFhYi0yZDdjZDAxMWRiNDciLCJ1bmlxdWVfbmFtZSI6ImdvZ29lQG1pY3Jvc29mdC5jb20iLCJ1cG4iOiJnb2dvZUBtaWNyb3NvZnQuY29tIiwidXRpIjoiYnpWQ3pTVU1URUdWc1lTdGhBRUNBQSIsInZlciI6IjEuMCJ9.S5BzTXlzK6R9j7l7eXqkJWOo5cwMCVKTk3NrAa2J3JTNrnHw40jHLfdm2tsayAmXgUoAr2UznV4nCR9GvWqdXlJdEJ6G4IG5TOpILbOgkBtYJk8gD-RUieRvQsGZoR4BzQx8fItHbMUkcpsv8y8W44Z_5fWmEQXjlOj0mspt6TeYLoDqbPHJbVG2WCD5RoX0roE9lWaxGVP0uYOCWc4GsFaCOwloXQd9uNrei2-uBO8qzdWadhcpe_c-z6EYlpWlJfHqzd4vMAixug5T-C6BKougELSDvLksFoiIMamBzOp1hTbWqtW_IPo4O8QDoXpBe-ZnXmWonZTJQZnLTqTcQQ'
  }
}).then(res => res.json())
.catch(error => console.error('Error:', error))
.then(response => {
  // console.log('Success:', response);
  // console.log(response.value[2]);
  // for(let i=0; i < response.value.length;i++){
  //   items.push({
  //   key: 'item' + i,
  //   name: "hello",
  //   position: '04 July, 2018',
  //   office: randWord(positions),   
  //   another: '10 KB',
  //   index: i
  // });
  // }

  
});
             


console.log(items);

class App extends React.Component {
  constructor(props){
    super(props);
    for (let i = 0; i < 100; i++) { 
      items.push({
        key: 'item' + i,
        name: 'Item ' ,
        position: '04 July, 2018',
        office: randWord(positions),
        another: '10 KB',
        index: i
      });
    }
  }
  public render() {
    return (
        <div>
          <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.0.0/css/fabric.min.css"/>
          <h1 className="ComponentPage-subHeading">Files</h1>
          <DetailsList items={ items } columns={ columns } selectionMode= {SelectionMode.none}/>
        </div>
    );
  }
}






export default App;





// ReactDOM.render(<MyPage />, document.body.firstChild);