import './App.css';
import * as React from 'react';
import {
  DocumentCard,
  DocumentCardActivity,
  DocumentCardPreview,
  DocumentCardTitle
} from 'office-ui-fabric-react/lib/DocumentCard';
import { initializeIcons } from '@uifabric/icons';

import {
  Selection,
} from 'office-ui-fabric-react/lib/DetailsList';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { adalApiFetch } from './adalConfig';
import { Breadcrumb } from 'office-ui-fabric-react/lib/Breadcrumb';
import { fileIcons } from './fileicons';
import { ItemsList } from './components/ItemsList';
import { Head } from './components/Head';
// Register icons and pull the fonts from the default SharePoint cdn:
initializeIcons();


const columns = [
  { 
    key: "icon", 
    name: "...", 
    fieldName: "icon", 
    minWidth: 20, 
    maxWidth: 20,
    isResizable: false,  
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
    key: "modified", 
    name: "Modified",
    fieldName: "modified",
    minWidth: 20, 
    maxWidth: 300,
    isResizable: true
  },
  { 
    key: "modifiedBy", 
    name: "Modified By",
    fieldName: "modifiedBy",
    minWidth: 160, 
    maxWidth: 280,
    isResizable: true       
  },
  {
    key: "size",
    name: 'Size',
    fieldName: "size",
    minWidth: 70,
    maxWidth: 100
  }
];

interface Users {  
    id: object;
}
 
class App extends React.Component<{},{items: Array<any>, folders: Array<any>, users: Users}> {
  public _selection: Selection;

  updateNavList: any = (obj) => {
    this.setState((prevState) => {
      let newList = prevState.folders; 
      let found = false;
      for(let i=0;i<prevState.folders.length;i++){
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
    let newObj = { text: 'Files', key: 'root', onClick: this.onBreadcrumbItemClicked };
    newObj.text = obj.name;
    newObj.key = obj.path + "/" + obj.name;
    this.updateNavList(newObj);
  }

  onBreadcrumbItemClicked: any = (ev: React.MouseEvent<HTMLElement>, item: any) => {
    let url = '';
    if(item.key == 'root') {
      url = 'https://graph.microsoft.com/v1.0/me/drive/root/children';
    }
    else{
      url = 'https://graph.microsoft.com/v1.0/me' + item.key + ":/children";
    }
    this.fetchFromDrive(url);
    this.updateNavList(item);
  }

  fetchFromDrive: any = (url) => {
    adalApiFetch(fetch, url, {}).then((response) => {
        response.json().then((response) => {
            this.setState((prevState) => ({
              items: []         
            }));
            let users = [];
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
                modified: d,
                modifiedBy: value.lastModifiedBy.user.displayName,
                size: size,
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
              let userId = value.lastModifiedBy.user.id;
              if((users.indexOf(userId) == -1)){
                users.push(userId);
              }
            }
            for(let i=0;i<users.length;i++){
              let id = users[i];
              let url = "https://graph.microsoft.com/v1.0/users/" + id;
              adalApiFetch(fetch, url, {}).then((response) => {
                response.json().then((response) => {
                    let photourl = url + "/photo/$value";
                    adalApiFetch(fetch, photourl, {})
                      .then((res) => (res.blob()))
                      .then((blob) => {
                         let urlCreator = window.URL;
                         let imageUrl = urlCreator.createObjectURL(blob);
                         response.imageUrl = imageUrl;
                        this.setState((prevState)=>{
                          let newUsers: any = prevState.users;
                          newUsers[id] = response;
                          return {users: newUsers}
                        });
                      })
                      .catch((error) => {
                        console.error(error);
                      });
                });
              }).catch((error) => {
                console.error(error);
              });
            }
        });
      }).catch((error) => {
        console.error(error);
      });
  }

  constructor(props){
    super(props);
    this.state = {
      items: [],
      folders: [{ text: 'Files', key: 'root', onClick: this.onBreadcrumbItemClicked }],
      users: {
        id: {}
      }
    };
    this._selection = new Selection({
      onSelectionChanged: () => {
        let obj = this._selection.getSelection()[0] as any;
        if(obj != undefined) {
          if('folder' in obj.value) this.updateList(obj);
        }
      }
    });
    this.fetchFromDrive('https://graph.microsoft.com/v1.0/me/drive/root/children');
  }

  public render() {
    return (
        <div>
          <Head />
          <Breadcrumb
          items={this.state.folders}
          />
          <ItemsList 
            items={ this.state.items }
            columns={ columns }
            selection={this._selection}
            users={this.state.users}
          />
        </div>
    );
  }
}

export default App;