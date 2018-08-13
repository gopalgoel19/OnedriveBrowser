
import * as React from 'react';
import {
  DetailsList,
  Selection,
  SelectionMode
} from 'office-ui-fabric-react/lib/DetailsList';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { MyHoverCard } from './MyHoverCard';


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
        if(item.type === 'file') return <i className="ms-Icon ms-Icon--FileTemplate" aria-hidden="true"></i>
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
      if(item.type === 'folder'){
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
    if (column.key === 'modifiedBy') {
      return (
        <MyHoverCard item={item} users={this.props.users}/>
      );
    }
    return item[column.key];
  };
}