import './App.css';
import * as React from 'react';
import { initializeIcons } from '@uifabric/icons';
import {
  Selection,
} from 'office-ui-fabric-react/lib/DetailsList';
import { adalApiFetch } from './adalConfig';
import { Breadcrumb } from 'office-ui-fabric-react/lib/Breadcrumb';
import { ItemsList } from './components/ItemsList';
import { Head } from './components/Head';
import { Query, withApollo } from "react-apollo";
import gql from "graphql-tag";

// Register icons and pull the fonts from the default SharePoint cdn:
initializeIcons();

interface Users {  
    id: object;
}

// const users_query = gql`
//   {
//     user
//   }
// `;

// const Users = (props) => { 
//   return (
//   <Query
//     query={users_query}
//     variables = {{
//       url: props.url
//     }}
//     >
//     {
//       ({loading,error,data})=>{
//         if(loading) return <div>Loading...</div>;
//         if(error) return <div>Error :</div>;
//         return <div>
//           hello
//         </div>
//       }
//     }
//   </Query>
// )};

class App extends React.Component<{},{items: Array<any>, folders: Array<any>, users: Users}> {
  public _selection: Selection;
  private client: any;

  loadNewFolderData: any = (item) => {
    this.fetchItemsFromOneDrive('https://graph.microsoft.com/v1.0/me' + item.path + "/" + item.name + ":/children");
    let newBreadcrumbObj = { 
      text: item.name, 
      key: item.path + "/" + item.name, 
      onClick: this.onbreadcrumbObjClicked 
    };
    this.updateBreadCrumbList(newBreadcrumbObj);
  }

  fetchItemsFromOneDrive: any = (url) => {
    const items_query = gql`
      query GetItems($url: String){
        items(url: $url){
          icon
          id
          name
          modified
          modifiedBy
          modifiedByUserId
          size
          url
          path
          type
        }
      }
    `;

    this.client.query({
      query: items_query,
      variables: {
        url: url
      }
    })
    .then((response)=>{
      this.setState(() => ({
        items: []         
      }));
      let users = [];
      const items = response.data.items;
      for (let i = 0; i < items.length; i++) { 
        let item = items[i];              
        this.pushItemToStateItemsList(item);

        let userId = item.modifiedByUserId;
        if((users.indexOf(userId) == -1)){
          users.push(userId);
        }
      }

      this.fetchUsersDataFromOneDrive(users);
    });
  }

  updateBreadCrumbList: any = (breadcrumbObj) => {
    this.setState((prevState) => {
      let newList = prevState.folders; 
      let found = false;
      for(let i=0;i<prevState.folders.length;i++){
        if( prevState.folders[i].key === breadcrumbObj.key){
          newList = newList.slice(0,i+1);
          found = true;
          break;
        } 
      }
      if(!found){
        newList = newList.concat(breadcrumbObj)
      }
      return {folders: newList};
    });
  }

  pushItemToStateItemsList: any = (item) => {
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

  onbreadcrumbObjClicked: any = (ev: React.MouseEvent<HTMLElement>, breadcrumbObj: any) => {
    let url = '';
    if(breadcrumbObj.key == 'root') {
      url = 'https://graph.microsoft.com/v1.0/me/drive/root/children';
    }
    else{
      url = 'https://graph.microsoft.com/v1.0/me' + breadcrumbObj.key + ":/children";
    }
    this.fetchItemsFromOneDrive(url);
    this.updateBreadCrumbList(breadcrumbObj);
  }

  constructor(props){
    super(props);
    this.state = {
      items: [],
      folders: [{ text: 'Files', key: 'root', onClick: this.onbreadcrumbObjClicked }],
      users: {
        id: {}
      }
    };
    this._selection = new Selection({
      onSelectionChanged: () => {
        let item = this._selection.getSelection()[0] as any;
        if(item != undefined) {
          if(item.type === "folder") this.loadNewFolderData(item);
        }
      }
    });
    this.client = props.client;
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

export default withApollo(App);