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
import { Breadcrumb, IBreadcrumbItem, IDividerAsProps } from 'office-ui-fabric-react/lib/Breadcrumb';
// Register icons and pull the fonts from the default SharePoint cdn:
initializeIcons();


const columns = [
  { 
    key: "icon", 
    name: "...", 
    fieldName: "icon", 
    minWidth: 20, 
    maxWidth: 20,
    isResizable: true,  
    onRender: (item) => {
      if(item.icon == ''){
        if('file' in item.value) return <i className="ms-Icon ms-Icon--FileTemplate" aria-hidden="true"></i>
        else return <i className="ms-Icon ms-Icon--FabricFolderFill" aria-hidden="true"></i>
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
    onRender: (item) => {
      if('folder' in item.value){
        return <Link href='' className="ms-font-m" style={{textDecoration:'none', color: 'black'}}>{item.name}</Link>
      }
      else{
        return <Link href={item.url} target='_blank' className="ms-font-m" style={{textDecoration:'none', color: 'black'}}>{item.name}</Link>
      }
    }
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
    isResizable: true       
  },
  {
    key: "another",
    name: 'Size',
    fieldName: 'another',
    minWidth: 70,
    maxWidth: 100
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

class App extends React.Component<{},{items: Array<any>, folders: Array<any>}> {
  private _selection: Selection;

  updateNavList: any = (obj) => {
    this.setState((prevState) => {
      let newList = prevState.folders; 
      let found = false;
      for(let i=0;i<prevState.folders.length;i++){
        // console.log('finding...');
        if(Object.is(prevState.folders[i],obj)){
          newList = newList.slice(0,i+1);
          found = true;
          break;
        }
      }
      if(!found){
        newList = newList.concat(obj)
      }
      return {folders: newList};
    });
  }

  updateList: any = (obj) => {
    this.fetchFromDrive('https://graph.microsoft.com/v1.0/me' + obj.path + "/" + obj.name + ":/children");
    // console.log(obj);
    let newObj = { text: 'Files', key: 'root', onClick: this.onBreadcrumbItemClicked };
    newObj.text = obj.name;
    newObj.key = obj.path + "/" + obj.name;
    // console.log(newObj);
    this.updateNavList(newObj);
    // console.log(this.state.folders);
  }

  onBreadcrumbItemClicked: any = (ev: React.MouseEvent<HTMLElement>, item: any) => {
    // alert(`Breadcrumb item with key "${item.key}" has been clicked.`);
    // console.log(item);
    let url = '';
    if(item.key == 'root') {
      url = 'https://graph.microsoft.com/v1.0/me/drive/root/children';
    }
    else{
      url = 'https://graph.microsoft.com/v1.0/me' + item.key + ":/children";
    }
    this.fetchFromDrive(url);
    // console.log(item);
    this.updateNavList(item);
  }

  fetchFromDrive: any = (url) => {
    adalApiFetch(fetch, url, {}).then((response) => {
        response.json().then((response) => {
            // console.log(JSON.stringify(response, null, 2));
            this.setState((prevState) => ({
              items: []         
            }));

            for (let i = 0; i < response.value.length; i++) { 
              let value = response.value[i];
              let size: any = value.size/1024;
              let d: any = new Date(value.lastModifiedDateTime);
              d = d.toString().slice(0,25);
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
                url: '',
                path: '',
                value: value
              };

              let icon = '';
              let name = value.name;
              for (let i=0; i< fileIcons.length ; i++){
                if(name.endsWith('.' + fileIcons[i].name)){
                  icon = `https://static2.sharepointonline.com/files/fabric/assets/brand-icons/document/svg/${fileIcons[i].name}_16x1.svg`;
                }
              }
              item.icon = icon;
              item.url = value.webUrl;
              item.path = value.parentReference.path;
              this.setState((prevState) => ({
                items: prevState.items.concat(item)
              }));
            }
        });
      }).catch((error) => {
        console.error(error);
      });
  }

  navLinkUpdate: any = (props) => {
    console.log(props);
  }

  

  constructor(props){
    super(props);
    this.state = {
      items: [],
      folders: [{ text: 'Files', key: 'root', onClick: this.onBreadcrumbItemClicked }]
    };
    this._selection = new Selection({
      onSelectionChanged: () => {
        let obj = this._selection.getSelection()[0] as any;
        if(obj != undefined) {
          // console.log(obj);
          if('folder' in obj.value) this.updateList(obj);
        }
      }
    });
    this.fetchFromDrive('https://graph.microsoft.com/v1.0/me/drive/root/children');
  }

  public render() {
    return (
        <div>
          <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.0.0/css/fabric.min.css"/>
          <div className="ms-BrandIcon--icon96 ms-BrandIcon--onedrive"></div>
          <Breadcrumb
          items={this.state.folders}
          ariaLabel={'Website breadcrumb'}
          />
          <DetailsList 
            items={ this.state.items }
            columns={ columns }
            selectionMode= {SelectionMode.none}
            selection={this._selection}
          />
        </div>
    );
  }
}

export default App;
