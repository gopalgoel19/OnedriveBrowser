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
    onRender: (item) => {
      if(item.icon == ''){
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
  
  updateList: any = (obj) => {
    this.fetchFromDrive('https://graph.microsoft.com/v1.0/me' + obj.path + "/" + obj.name + ":/children");

  }

  _onBreadcrumbItemClicked: any = (ev: React.MouseEvent<HTMLElement>, item: any) => {
    alert(`Breadcrumb item with key "${item.key}" has been clicked.`);
  }

  fetchFromDrive: any = (url) =>{
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
                url: '',
                path: ''
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
    this.fetchFromDrive('https://graph.microsoft.com/v1.0/me/drive/root/children');
    this.state = {
      items: [],
      folders: []
    };

    this._selection = new Selection({
      onSelectionChanged: () => {
        let obj = this._selection.getSelection()[0] as any;
        if(obj != undefined) {
          console.log(obj);
          if(obj.icon == "") this.updateList(obj);
        }
      }
    });
  
  }

  public render() {
    return (
        <div>
          <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.0.0/css/fabric.min.css"/>
          <div className="ms-BrandIcon--icon96 ms-BrandIcon--onedrive"></div>
          <Breadcrumb
          items={[
            { text: 'Files', key: 'Files', onClick: this._onBreadcrumbItemClicked },
            { text: 'This is folder 1', key: 'f1', onClick: this._onBreadcrumbItemClicked },
            { text: 'This is folder 2', key: 'f2', onClick: this._onBreadcrumbItemClicked },
            { text: 'This is folder 3', key: 'f3', onClick: this._onBreadcrumbItemClicked },
            { text: 'This is folder 4', key: 'f4', onClick: this._onBreadcrumbItemClicked },
            { text: 'This is folder 5', key: 'f5', onClick: this._onBreadcrumbItemClicked, isCurrentItem: true }
          ]}
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
