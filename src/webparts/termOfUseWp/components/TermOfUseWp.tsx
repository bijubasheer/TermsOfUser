import * as React from 'react';
import styles from './TermOfUseWp.module.scss';
import { ITermOfUseWpProps } from './ITermOfUseWpProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/site-users";
import { Button, ButtonType, Panel, Modal,IconButton, IIconProps, Stack, IStackTokens } from 'office-ui-fabric-react';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { _RoleDefinition } from '@pnp/sp/security/types';

var content = "";
var title = "";
var email = '';
const stackTokens: IStackTokens = { childrenGap: 40 };
export interface ISidePanelState {
  isOpen?: boolean;
}

export interface IDialogcontrolState{  
  hideDialog: boolean;  
}  

export default class TermOfUseWp extends React.Component<ITermOfUseWpProps, IDialogcontrolState> {

  public constructor(props: ITermOfUseWpProps, state: ISidePanelState) {
    super(props, state);
    email = props.email;
    content = props.content;
    title = props.title;

    //this.LoadTermsContent();

    this.state = {  
      hideDialog: true
    };  
  }

  // private async LoadTermsContent()
  // {
  //   await sp.web.lists.getByTitle("Terms of Use").items.select("Title", "Content").top(1).orderBy("Modified", true).get().then(items =>
  //     {
  //       //console.log(items);
  //       content = items[0]["Content"];
  //       title = items[0]["Title"];
  //     }
  //   );
  // }
  public render(): React.ReactElement<ITermOfUseWpProps> {    
    
    return (
      <>
      
     <div style={{margin:"10px"}}>
       <div>
        Terms of Use
      </div> 
      <Modal
        titleAriaId={title}
        isOpen={this.state.hideDialog}  
        onDismiss={this._accept}
        isBlocking={true}        
      >
        <div >
          <h2><span id={'blahId'}>{title}</span></h2>
          <IconButton
            ariaLabel="Close popup modal"
            onClick={_reject}
          />
        </div>
        <div dangerouslySetInnerHTML={{ __html: content}}/>
        <div className={'.center'} >

          <Stack horizontal tokens={stackTokens}>
            <DefaultButton text="Accept" onClick={this._accept} allowDisabledFocus disabled={false}  />
            <PrimaryButton text="Reject" onClick={_reject} allowDisabledFocus disabled={false}  />
          </Stack>
          <div><p></p></div>
        </div>
      </Modal>
      </div>
    </>
    );
    
    function _reject(): void {      
      window.location.href = '/sites/DealerDaily';      
    }
  }

  
  private _showDialog = (): void => {  
    this.setState({ hideDialog: false });  
  }  
  private toggleHideDialog() {
    this.setState({
      hideDialog: !this.state.hideDialog
    });
  }
  private _accept = (): void => {      
    this.Update();
    alert("Thanks for accpeting the Terms of Use.");
    this.setState({ hideDialog: false });  
  }  

  private async Update()
  {
    
    const items: any[] = await sp.web.lists.getByTitle("Terms of Use").items.top(1).orderBy("Modified", true).get();
    if (items.length > 0) {
      let currentAcceptedBy : string = items[0]["AcceptedBy"];

      console.log(currentAcceptedBy);
      if(currentAcceptedBy.indexOf(email) >= 0)
      {
        // const updatedItem = await sp.web.lists.getByTitle("Terms of Use").items.getById(items[0].Id).update({
        //   AcceptedBy: email
        // });
      }
      else
      {
        currentAcceptedBy = currentAcceptedBy + "," + email ;
        const updatedItem = await sp.web.lists.getByTitle("Terms of Use").items.getById(items[0].Id).update({
          AcceptedBy: currentAcceptedBy
        });
      }

  }

  
  }
}