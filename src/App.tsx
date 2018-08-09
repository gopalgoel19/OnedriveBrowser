import './App.css';
import * as React from 'react';
import {
  DocumentCard,
  DocumentCardActivity,
  DocumentCardPreview,
  DocumentCardTitle
} from 'office-ui-fabric-react/lib/DocumentCard';
import { initializeIcons } from '@uifabric/icons';
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import {
  DetailsList,
  Selection,
  SelectionMode
} from 'office-ui-fabric-react/lib/DetailsList';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { authContext } from './adalConfig';
import { adalApiFetch } from './adalConfig';
import { Breadcrumb } from 'office-ui-fabric-react/lib/Breadcrumb';
import { HoverCard, IExpandingCardProps } from 'office-ui-fabric-react/lib/HoverCard';
import { IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { fileIcons } from './fileicons';
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
  private _selection: Selection;

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

  navLinkUpdate: any = (props) => {
    console.log(props);
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

  logout: any = () => {
    authContext.logOut();
  }

  public render() {
    return (
        <div>
          <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.0.0/css/fabric.min.css"/>
          <div className="ms-BrandIcon--icon96 ms-BrandIcon--onedrive"></div>
          <DefaultButton
            data-automation-id="test"
            allowDisabledFocus={true}
            text="Logout"
            onClick={this.logout}
            style={{float:'right'}}
          />
          <Breadcrumb
          items={this.state.folders}
          ariaLabel={'Website breadcrumb'}
          />
          <DetailsList 
            items={ this.state.items }
            columns={ columns }
            selectionMode= {SelectionMode.none}
            selection={this._selection}
            onRenderItemColumn={this.onRenderItemColumn}
          />
        </div>
    );
  }

  private onRenderItemColumn: any = (item: any, index: number, column: IColumn) => {
    const expandingCardProps: IExpandingCardProps = {
      onRenderCompactCard: this._onRenderCompactCard,
      onRenderExpandedCard: this._onRenderExpandedCard,
      renderData: item
    };
    if (column.key == 'office') {
      return (
        <HoverCard id="myID1" expandingCardProps={expandingCardProps} instantOpenOnClick={true}>
          <div className="HoverCard-item" data-is-focusable={true}>
            {item.office}
          </div>
        </HoverCard>
      );
    }
    return item[column.key];
  };

  private _onRenderCompactCard = (item: any): JSX.Element => {
    let id = item.value.lastModifiedBy.user.id;
    let user = this.state.users[id];
    return (
      <div className="hoverCardExample-compactCard">

      <span style={{display: 'inline-block', width: '140px', height: 'auto'}}>
          <img aria-hidden="true" src={user.imageUrl}
          style={{display: 'inline', width: '100%', height: 'auto', padding: '10px', borderRadius: '50%'}}/>
      </span>
      <span style={{display: 'inline-block', padding: '0px'}} >
          
          <div className="hoverCardExample-expandedCard" style={{margin: '10px'}}>
            <div>
              <span className="ms-Icon ms-Icon--Contact" aria-hidden="true" style={{padding: '2px'}}></span><span>{user.displayName}</span>
            </div>
            <div>
              <span className="ms-Icon ms-Icon--Education" aria-hidden="true" style={{padding: '2px'}}></span><span>{user.jobTitle}</span>
            </div>     
          </div>
      </span>
      </div>
    );
  };

  private _onRenderExpandedCard = (item: any): JSX.Element => {
    let id = item.value.lastModifiedBy.user.id;
    let user = this.state.users[id];
    return (
      <div className="hoverCardExample-expandedCard" style={{margin: '10px'}}>
        <div className='ms-font-su'>Contact</div>
        <div style={{padding: '2px'}}>
          <span className="ms-Icon ms-Icon--MailSolid" aria-hidden="true" style={{padding: '5px'}}></span><span className='ms-font-m'> {user.mail}</span>
        </div>
        <div style={{padding: '2px'}}>
          <span className="ms-Icon ms-Icon--Location" aria-hidden="true" style={{padding: '5px'}}></span><span className='ms-font-m'> {user.officeLocation}</span>
        </div>
        <div style={{padding: '2px'}}>
          <span className="ms-Icon ms-Icon--Phone" aria-hidden="true" style={{padding: '5px'}}></span><span className='ms-font-m'> {user.businessPhones}</span>
        </div>      
      </div>
    );
  };
}

export default App;