
import * as React from 'react';
import { HoverCard, IExpandingCardProps } from 'office-ui-fabric-react/lib/HoverCard';

interface Users {  
  id: object;
}

export class MyHoverCard extends React.Component<{item: any, users: Users },{}> {
    
  public render() {
    const expandingCardProps: IExpandingCardProps = {
        onRenderCompactCard: this._onRenderCompactCard,
        onRenderExpandedCard: this._onRenderExpandedCard,
        renderData: this.props.item
      };
    return (
      <HoverCard id="myID1" expandingCardProps={expandingCardProps} instantOpenOnClick={true}>
        <div className="HoverCard-item" data-is-focusable={true}>
          {this.props.item.modifiedBy}
        </div>
      </HoverCard>
    );
  }

  private _onRenderCompactCard = (item: any): JSX.Element => {
    let id = item.modifiedByUserId;
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
    let id = item.modifiedByUserId;
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
