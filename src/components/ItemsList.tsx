
import * as React from 'react';
import {
  DocumentCard,
  DocumentCardActivity,
  DocumentCardPreview,
  DocumentCardTitle
} from 'office-ui-fabric-react/lib/DocumentCard';
import { initializeIcons } from '@uifabric/icons';
import {
  DetailsList,
  Selection,
  SelectionMode
} from 'office-ui-fabric-react/lib/DetailsList';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { HoverCard, IExpandingCardProps } from 'office-ui-fabric-react/lib/HoverCard';
import { IColumn } from 'office-ui-fabric-react/lib/DetailsList';
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

export class ItemsList extends React.Component<{items: Array<any>, selection: Selection, users: Users},{}> {

  public render() {
    return (
      <DetailsList 
        items={ this.props.items }
        columns={ columns }
        selectionMode= {SelectionMode.none}
        selection={this.props.selection}
        onRenderItemColumn={this.onRenderItemColumn}
      />
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
    let user = this.props.users[id];
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
    let user = this.props.users[id];
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

// export default ItemsList;