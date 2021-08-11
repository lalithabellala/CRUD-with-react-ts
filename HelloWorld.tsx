import * as React from 'react';
import styles from './HelloWorld.module.scss';
import { IHelloWorldProps } from './IHelloWorldProps';
import { escape } from '@microsoft/sp-lodash-subset';
import {TextField} from "office-ui-fabric-react";//importing TextField from office-ui-fabric
import { Checkbox, ICheckboxProps } from '@fluentui/react/lib/Checkbox';
import { ChoiceGroup, IChoiceGroupOption } from '@fluentui/react/lib/ChoiceGroup';
import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { IStackTokens, Stack } from '@fluentui/react/lib/Stack';
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { sp,Web, IWeb } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/views";
import "@pnp/sp/items";
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import pnp from 'sp-pnp-js';
//People picker
import { IPersonaProps } from '@fluentui/react/lib/Persona';
import { IBasePickerSuggestionsProps, NormalPeoplePicker, ValidationState } from '@fluentui/react/lib/Pickers';

import {PeoplePicker,PrincipalType} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { peoplePicker } from '@fluentui/react/lib/components/ExtendedPicker/PeoplePicker/ExtendedPeoplePicker.scss';
//import { autobind } from 'office-ui-fabric-react/lib/Utilities';
// import pnp, { Web } from 'sp-pnp-js';// importing sp keyword form pnp-js

//for radio button
const options: IChoiceGroupOption[] = [
  { key: 'Open', text: 'Open' },
  { key: 'Closed', text: 'Closed' },];

//for dropdowns
const drops: IDropdownOption[] = [
  { key: 'JNJ specific', text: 'Select an option' , itemType: DropdownMenuItemType.Header},
  { key: 'JNJ specific', text: 'JNJ specific?' },
  { key: 'Jabil Specific', text: 'Jabil Specific?' },
  { key: 'CR Specific', text: 'CR Specific?' },
  { key: ' CN Specific', text: ' CN Specific?' },
  
];


const suggestionProps: IBasePickerSuggestionsProps = {
  suggestionsHeaderText: 'Suggested People',
  mostRecentlyUsedHeaderText: 'Suggested Contacts',
  noResultsFoundText: 'No results found',
  loadingText: 'Loading',
  showRemoveButtons: true,
  suggestionsAvailableAlertText: 'People Picker Suggestions available',
  suggestionsContainerAriaLabel: 'Suggested contacts',
};
//let Title:[];

//for edit 
var hilIdItems: IDropdownOption[]=[];

let ppl:any[]=[];

var hilId;
export interface IStates {
  cR:any
    cN:any,
  changeOwn:any,
    changeOwnPh:any,
    pLM:boolean,
    changeSta:any,
    jCR:any,
    jCN:any,
    jchangeOwn:any,
    jchangeOwnph:any,
    notificationType:any,
    notificationDesc:any,
    notificationDetail:any,
    returnEvi:boolean,
    notificationCompletion:any,
    eviReceive:boolean,
    //id based retrieval
    hilIdItems: any;
    hilId: any;
    hidehilid:boolean;
    //for people picker
   
 // UserDetails: IUserDetail[];
  selectedusers: any[];
 


   
}


export default class HelloWorld extends React.Component<IHelloWorldProps, IStates> {


  constructor(props)
  {
    super(props);
    sp.setup({
      spfxContext: this.props.context
    });
   
    //this.getPeoplePicker=this.getPeoplePicker.bind(this);
    this.state={
      
    cR:'',
    cN:'',
    pLM:false,
    changeOwn:'',
    changeOwnPh:'',
   changeSta:'',
    jCR:'',
    jCN:'',
    jchangeOwn:'',
    jchangeOwnph:'',
    notificationType:'',
    notificationDesc:'',
    notificationDetail:"",
    returnEvi:false,
    notificationCompletion:'',
    eviReceive:false,
   //radiobutton
  //  changeSta:"",
  //  singleValueOptions:[],
   //dropdown
  //  notificationType:""
      hilIdItems: [],

   hilId:'',
   hidehilid:false,

   
   // UserDetails: IUserDetail[];
    selectedusers: [],
      
    
    }
    this.handleChangeId = this.handleChangeId.bind(this);
    this.getJnjowner=this.getJnjowner.bind(this);
  }

  // @autobind
  // private _getPeoplePickerItems(items: any[]) {
  //   let userarr: IUserDetail[] = [];
  //   items.forEach(user => {
  //     userarr.push({ ID: user.id, LoginName: user.loginName });
  //   })
  //   this.setState({ UserDetails: userarr })
  // }
  async handleChangeId(event): Promise<void> {
    try {
    hilId = event.key;
    let items = await pnp.sp.web.lists.getByTitle("NotificationFormField").items.getPaged();
    for (let i = 0; i < items.results.length; i++) {
     
    if (items.results[i].ID == hilId) 
    {
    this.setState({ cR:items.results[i].J_x0026_J_x0020_CR_x002f_DCR_x00 });
    this.setState({ cN:items.results[i].  J_x0026_J_x0020_CN_x002f_DCO});
    this.setState({ pLM:items.results[i].  J_x0026_J_x0020_PLM});
    this.setState({ notificationType:items.results[i].Notification_x0020_Type  });
    this.setState({ changeOwn:items.results[i].  J_x0026_J_x0020_Change_x0020_Own });
    this.setState({ changeOwnPh:items.results[i].  J_x0026_J_x0020_Change_x0020_Own0 });
    this.setState({ changeSta:items.results[i].  Change_x0020_Status });
    this.setState({ jCR:items.results[i]. Jabil_x0020_CR });
    this.setState({ jCN:items.results[i]. Jabil_x0020_CN});
    this.setState({ jchangeOwn:items.results[i]. Jabil_x0020_Change_x0020_Owner});
    this.setState({ jchangeOwnph:items.results[i].  Jabil_x0020_Change_x0020_Owner_x});
    this.setState({ notificationDesc:items.results[i]. Notification_x0020_Description});
    this.setState({ returnEvi:items.results[i].  Return_x0020_evidence_x0020_need});
    this.setState({ eviReceive:items.results[i].   Evidence_x0020_received});
    this.setState({selectedusers:items.results[i]. PeoplePicker})

  //....................................
      let dummyElement = document.createElement("DIV");
      dummyElement .innerHTML = items.results[i].Notification_x0020_Detail;

    this.setState({notificationDetail : dummyElement.innerText});
   
    dummyElement .innerHTML = items.results[i].Notification_x0020_Completion;
    this.setState({ notificationCompletion:dummyElement.innerText});
      
  
          
          }
        }
      } catch (error) {
    console.error(error);
    alert(error);
      }
     
    }

    async editToList(event): Promise<void> {
 
      try {
     
      this.setState({ hidehilid: true })
      let list = pnp.sp.web.lists.getByTitle("NotificationFormField");
      const i = await list.items.getById(hilId).update({
        J_x0026_J_x0020_CR_x002f_DCR_x00 :this.state.cR,
        J_x0026_J_x0020_CN_x002f_DCO:this.state.cN,
        J_x0026_J_x0020_Change_x0020_Own:this.state.changeOwn,
        J_x0026_J_x0020_Change_x0020_Own0:this.state.changeOwnPh,
        //for radio button
        Change_x0020_Status:this.state.changeSta,
        //.........fro dropdown
        Notification_x0020_Type:this.state.notificationType,
        //..........
        Jabil_x0020_CR:this.state.jCR,
        Jabil_x0020_CN:this.state.jCN,
        Jabil_x0020_Change_x0020_Owner:this.state.jchangeOwn,
        Jabil_x0020_Change_x0020_Owner_x:this.state.jchangeOwnph,
        Notification_x0020_Description:this.state.notificationDesc,
        Notification_x0020_Detail:this.state.notificationDetail,
       Return_x0020_evidence_x0020_need:this.state.returnEvi,
        Notification_x0020_Completion:this.state.notificationCompletion,
       Evidence_x0020_received:this.state.eviReceive,

       //........
       PeoplePickerId: this.state.selectedusers


    
              });
      alert("It is updated successfully in the list");
      this.clearFields();
            } catch (error) {
      console.error(error);
            }
          }
     
  public async componentDidMount(): Promise<void>
      {
        //alert("did mount method");
    // get all the items from a sharepoint list
    var reacthandler=this;
   // alert(hilIdItems);
    pnp.sp.web.lists.getByTitle("NotificationFormField").items.select('ID').get().then(function (data) {
    for (var k in data) {
    hilIdItems.push({ key:data[k].ID, text:data[k].ID });
          }
    reacthandler.setState({ hilIdItems:hilIdItems });
    console.log(hilIdItems);
   // alert(hilIdItems);
    return hilIdItems;
        });
        
      
       
      }
    
   
       
  private _alertClicked=async()=>{
   alert('Clicked');
 
    let status=this.state.changeSta;
    if(status==""){
    alert("Please Give the Change Status Value");}
    else{
      try{      
        
        let web=Web(this.props.webURL)
        alert("weburl is good");
        // alert(this.state.cR);
        
        await web.lists.getByTitle("NotificationFormField").items.add({
           
         
    //..........................
              J_x0026_J_x0020_CR_x002f_DCR_x00:this.state.cR,
              J_x0026_J_x0020_CN_x002f_DCO:this.state.cN,
              J_x0026_J_x0020_PLM:this.state.pLM,
              J_x0026_J_x0020_Change_x0020_Own:this.state.changeOwn,
              J_x0026_J_x0020_Change_x0020_Own0:this.state.changeOwnPh,
              //for radio button
              Change_x0020_Status:this.state.changeSta,
              //.........fro dropdown
              Notification_x0020_Type:this.state.notificationType,
              //..........
              Jabil_x0020_CR:this.state.jCR,
              Jabil_x0020_CN:this.state.jCN,
              Jabil_x0020_Change_x0020_Owner:this.state.jchangeOwn,
              Jabil_x0020_Change_x0020_Owner_x:this.state.jchangeOwnph,
              Notification_x0020_Description:this.state.notificationDesc,
              Notification_x0020_Detail:this.state.notificationDetail,
             Return_x0020_evidence_x0020_need:this.state.returnEvi,
              Notification_x0020_Completion:this.state.notificationCompletion,
             Evidence_x0020_received:this.state.eviReceive,

             //................need to add 'Id' at end to save the value
             PeoplePickerId: this.state.selectedusers
            
         })
     alert(this.state.selectedusers);
      alert("succesfull");
      this.clearFields();
               
    }
    catch(error)
    {
     
      console.error();
      alert(error+" in catch block");
    }
 }

  }

  public getPeoplePicker(items) {
    console.log(items);
    
    }

  //for resetting values
  private clearFields = (): void => {
   
    this.setState({ 
     cN:'',
     cR:'',
     pLM:false,
     changeOwn:'',
     changeOwnPh:'',
    changeSta:'',
     jCR:'',
     jCN:'',
     jchangeOwn:'',
     jchangeOwnph:'',
     notificationType:'',
     notificationDesc:'',
     notificationDetail:"",
     returnEvi:false,
     notificationCompletion:'',
     eviReceive:false,
     hidehilid:true
          
    });
  }


  //for people picker
  public getJnjowner(items:any[]) {
    console.log('Items:', items);  
  
  //   try{

     
  //   items.map((item)=>{
  //     ppl.push(item.secondaryText);
  //   });
  //   this.setState({ selectedusers:ppl });
  //   alert(ppl+"gfdg");
  //   alert(this.state.selectedusers+"jhbjhb");
  //  }
  //  catch(error)
  //  {
  //    alert(error);
  //  }

  //  this.setState({selectedusers:ppl});
 
  //  alert(ppl);
  //  ppl[0]=items[0].secondaryText;
  //  ppl[1]=items[1].secondaryText;

   this.setState({ selectedusers:items[0].id });
   alert(this.state.selectedusers);
//     try{
   
//    // this.setState({pplControl:items});
//   // 
//   items.map((item)=>{
//       ppl.push(item.id);
//        });
    
  
//   alert(ppl+"ppl");
//   // this.setState({selectedusers:ppl})
//   alert(ppl);
// }
// catch(error){
//   alert(error+"catch block error");
  
// }
  
    //
    // var i=0;
    // var len=items.length;
    // while(i==len)
    // {
    //   alert("inside loop"+i);
    //   ppl.push(items[i].id);//or use email
    //   ppl.push(items[i].secondaryText);
    //   alert(i+"inc");
    //   i++;
    // }
  //   let ppl:any[]=[];
     
  //   items.map((item)=>{
  //     ppl.push(item.secondaryText);
  //   });
  // //  alert(i+"loop out");
  //   alert(items.length+"lenght of items");
  //   alert(items);
  //   alert(ppl+"    ppl");
  //worked one............................................................
// this.setState({ selectedusers:items[0].id });
//...................................................................
//this.setState({ selectedusers:ppl });
  
    }


  
   
    public render(): React.ReactElement<IHelloWorldProps> {
    const dropdownStyles: Partial<IDropdownStyles> = {
      dropdown: { width: 300 },
    };
    const stackTokens: IStackTokens = { childrenGap: 20 };
    pnp.setup({
      spfxContext:this.props.context
    });
    return (
      <div className={ styles.helloWorld }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              {/* <PeoplePicker context={this.props.context}></PeoplePicker> */}
           
             <TextField className={styles.styling}
                type="text"  
                label="J&J CN/DCR"
                value ={this.state.cR}
                onChange={(event, value) => { this.setState({ cR: value }); }} >

                </TextField>

             <TextField 
             type="text"
             label="J&J CN/DCO"
             value={this.state.cN}
            
             onChange={(event, value) => { this.setState({ cN: value }); }}>

             </TextField>
            
             <br></br>
           
           <Checkbox 
                label='J&J PLM'

                checked={this.state.pLM}
                onChange={(event, value) => { this.setState({ pLM: value }); }}
                
           />
           
           
           
             <TextField
                label='J&J Change Owner '
                value ={this.state.changeOwn}
                onChange={(event, value) => { this.setState({ changeOwn: value }); }}
           />
            <TextField
                            
                label='J&J Change Owner phone' 
                value ={this.state.changeOwnPh}
                onChange={(event, value) => { this.setState({ changeOwnPh: value }); }}
           />

           <ChoiceGroup
                label='Change Status'
                 options={options} 
                 required
                 
                  onChange={(event,options)=>this.setState({ changeSta: options.key })} 
             />
          
  
        
            <TextField
                label='Jabil CR ' 
                value ={this.state.jCR}
                onChange={(event, value) => { this.setState({ jCR: value }); }}
           />
            <TextField
                label='Jabil CN  '
                value ={this.state.jCN}
                onChange={(event, value) => { this.setState({ jCN: value }); }}
           />


             <TextField
                label='Jabil Change Owner  '
                value={this.state.jchangeOwn}
                onChange={(event, value) => { this.setState({ jchangeOwn: value }); }}
           />

           
                
                <PeoplePicker 
                    context={this.props.context}
                    placeholder="Enter your name"
                    //personSelectionLimit={3}
                    groupName={"Jnj"}//it is used for filtering users from all users 
                    showtooltip={false}                   
                    disabled={false}
                    showHiddenInUI={false}
                    ensureUser={true}
                   required={true}
                   // selectedItems={this.getJnjowner}
                   // principalTypes={[PrincipalType.SharePointGroup]}//later we can modify it to group
                    resolveDelay={1000}
                   onChange={this.getJnjowner}
                    ></PeoplePicker>
                 
                  
           <TextField
                label='Jabil Change Owner phone '
                value={this.state.jchangeOwnph}
                onChange={(event, value) => { this.setState({ jchangeOwnph: value }); }}
           />
          <Stack tokens={stackTokens}>
         
           <Dropdown
            placeholder="Select an option"
            label="Notification Type:"
            options={drops}
            styles={dropdownStyles}
             onChange={(event,drops)=>{this.setState({ notificationType:drops.key});}}
            /> 
            </Stack>
          
      <TextField
           label=' Notification Description  '
           value={this.state.notificationDesc}
         
          onChange={(event,value) =>{this.setState({notificationDesc:value});}}
       />
        <TextField
           label=' Notification Detail  '
           value={this.state.notificationDetail}
           onChange={(event, value) => { this.setState({ notificationDetail: value }); }}
       /> <br></br>
       
       <Checkbox 
                label='Return evidence needed?'
               
                checked={this.state.returnEvi}
                // onChange={ this._onControlledCheckboxChange }
                onChange={(event, value) => { this.setState({ returnEvi: value }); }}
                
           />
                
           <br></br>
       
       <TextField
           label='Notification Completion'
           value={this.state.notificationCompletion}
           onChange={(event, value) => { this.setState({ notificationCompletion: value }); }}
       />
       <br></br>
       
       <Checkbox 
                label='Evidence received '
                checked={this.state.eviReceive}
              
                // onChange={ this._onControlledCheckboxChange }
                onChange={(event, value) => { this.setState({ eviReceive: value }); }}
                
           />
        <Dropdown 
                       placeholder="Select an ID to Edit "
                        label="Select an ID to Edit"
                          selectedKey={hilId}
                         options={this.state.hilIdItems}
                         styles={dropdownStyles}
                            onChanged={this.handleChangeId}
                            /> 
                           
                          
       <br></br>
       <Stack horizontal tokens={stackTokens}>
      
      <PrimaryButton text="Submit"  onClick={this._alertClicked}/><br></br><br></br><br></br>
       <PrimaryButton onClick={() => this.editToList(event)}>Edit</PrimaryButton>&nbsp;&nbsp; 
    </Stack> 
    
    </div> 
            
            </div>
          </div>
        </div>
      
    );
  }
}
