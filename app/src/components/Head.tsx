import * as React from 'react';
import { authContext } from './../adalConfig';
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
export class Head extends React.Component<{},{}>{
    
    logout: any = () => {
      authContext.logOut();
    }

    render(){
        return(
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
            </div>
        )
    }
}