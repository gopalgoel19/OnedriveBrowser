import './App.css';
import * as React from 'react';
import { initializeIcons } from '@uifabric/icons';
import {
  Selection,
} from 'office-ui-fabric-react/lib/DetailsList';
import { adalApiFetch } from './adalConfig';
import { Breadcrumb } from 'office-ui-fabric-react/lib/Breadcrumb';
import { fileIcons } from './fileicons';
import { ItemsList } from './components/ItemsList';
import { Head } from './components/Head';
// Register icons and pull the fonts from the default SharePoint cdn:
initializeIcons();

interface Users {  
    id: object;
}

let getSizeAsString = (sizeInBytes) => {
  let size: any = sizeInBytes/1024;
  if(size<1024){
    size = size.toFixed(2) + " KB";
    return size;
  }
  size = size/1024;
  if(size<1024){
    size = size.toFixed(2) + " MB";
    return size;
  }
  size = size/1024;
  size = size.toFixed(2) + " GB";
  return size;
};
 
class App extends React.Component<{},{items: Array<any>, folders: Array<any>, users: Users}> {
  public _selection: Selection;

  loadNewFolderData: any = (obj) => {
    this.fetchItemsFromOneDrive('https://graph.microsoft.com/v1.0/me' + obj.path + "/" + obj.name + ":/children");
    this.updateBreadCrumbList(obj);
  }
  
  fetchItemsFromOneDrive: any = (url) => {
    adalApiFetch(fetch, url, {}).then((response) => {
        response.json().then((response) => {
            this.setState((prevState) => ({
              items: []         
            }));
            let users = [];

            for (let i = 0; i < response.value.length; i++) { 
              let value = response.value[i];              
              this.pushItemToStateItemsList(value);

              let userId = value.lastModifiedBy.user.id;
              if((users.indexOf(userId) == -1)){
                users.push(userId);
              }
            }

            this.fetchUsersDataFromOneDrive(users);
        });
      }).catch((error) => {
        console.error(error);
      });
  }

  updateBreadCrumbList: any = (obj) => {
    let newBreadcrumbObj = { text: 'Files', key: 'root', onClick: this.onBreadcrumbItemClicked };
    if(obj.name) newBreadcrumbObj.text = obj.name;
    if(obj.path && obj.name) newBreadcrumbObj.key = obj.path + "/" + obj.name;
    this.setState((prevState) => {
      let newList = prevState.folders; 
      let found = false;
      for(let i=0;i<prevState.folders.length;i++){
        if( prevState.folders[i].key == newBreadcrumbObj.key){
          newList = newList.slice(0,i+1);
          found = true;
          break;
        } 
      }
      if(!found){
        newList = newList.concat(newBreadcrumbObj)
      }
      return {folders: newList};
    });
  }

  pushItemToStateItemsList: any = (value) => {
    let item = {
      icon: '',
      key: value.name,
      name: value.name,
      modified: new Date(value.lastModifiedDateTime).toString().slice(0,25),
      modifiedBy: value.lastModifiedBy.user.displayName,
      modifiedByUserId: value.lastModifiedBy.user.id,
      size: getSizeAsString(value.size),
      url: value.webUrl,
      path: value.parentReference.path,
      type: 'file' in value ? "file" : "folder"
    };
    let icon = '';
    let name = value.name;
    for (let i=0; i< fileIcons.length ; i++){
      if(name.endsWith('.' + fileIcons[i].name)){
        icon = `https://static2.sharepointonline.com/files/fabric/assets/brand-icons/document/svg/${fileIcons[i].name}_16x1.svg`;
      }
    }
    item.icon = icon;

    this.setState((prevState) => ({
      items: prevState.items.concat(item)
    }));  
  }

  fetchUsersDataFromOneDrive: any = (users) => {
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
  }

  onBreadcrumbItemClicked: any = (ev: React.MouseEvent<HTMLElement>, item: any) => {
    let url = '';
    if(item.key == 'root') {
      url = 'https://graph.microsoft.com/v1.0/me/drive/root/children';
    }
    else{
      url = 'https://graph.microsoft.com/v1.0/me' + item.key + ":/children";
    }
    this.fetchItemsFromOneDrive(url);
    this.updateBreadCrumbList(item);
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
        // console.log(obj);
        if(obj != undefined) {
          if(obj.type === "folder") this.loadNewFolderData(obj);
        }
      }
    });
    this.fetchItemsFromOneDrive('https://graph.microsoft.com/v1.0/me/drive/root/children');
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
            selection={this._selection}
            users={this.state.users}
          />
        </div>
    );
  }
}

export default App;