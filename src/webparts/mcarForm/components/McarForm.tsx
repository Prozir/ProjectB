import * as React from 'react';
import styles from './McarForm.module.scss';
import { IMcarFormProps } from './IMcarFormProps';
import { IListItem } from './IListItem';
import { escape } from '@microsoft/sp-lodash-subset';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { Dropdown, DropdownMenuItemType, IDropdownOption, IDropdownProps } from 'office-ui-fabric-react/lib/Dropdown';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { DatePicker } from 'office-ui-fabric-react/lib/DatePicker';
import { FontSizes } from 'office-ui-fabric-react/lib/Styling';
import { Image } from 'office-ui-fabric-react/lib/Image';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { DefaultButton, PrimaryButton, Stack, IStackTokens, setMemoizeWeakMap } from 'office-ui-fabric-react';
import { TooltipHost, ITooltipHostStyles } from 'office-ui-fabric-react/lib/Tooltip';
import {
  Accordion,
  AccordionItem,
  AccordionItemHeading,
  AccordionItemButton,
  AccordionItemPanel,
} from 'react-accessible-accordion';
import 'react-accessible-accordion/dist/fancy-example.css';
import * as jQuery from "jquery";

export interface IMcarFormState {
  mode?: string | null;
  isexistingsolution?: boolean | false;
  issharepointrelated?: boolean | false;
 // carouselElements: JSX.Element[];
 items:IListItem[];
 valuesobj:any;
}
let mcaritemid="";
var newitemobj={};
let ischecked:boolean=false;
var edititemobj={};
let getSelectedO365Keys = [];
let getSelectedTechUsedKeys=[];
let getSelectedAuthModels=[];
let requiredfields=["Solution_x0020_Name","Line_x0020_of_x0020_Business","Solution_x0020_OwnerId","Solution_x0020_SponsorId","Description","Business_x0020_case","With_x0020_which_x0020_O365_x002",
"Technology_x0020_used","Solution_x0020_type","If_x0020_yes_x002c__x0020_please","If_x0020_yes_x002c__x0020_list_x"];
const optionsspcarid: IDropdownOption[] = [
  { key: 'fruitsHeader', text: 'Fruits', itemType: DropdownMenuItemType.Header },
  { key: 'apple', text: 'Apple' },
  { key: 'banana', text: 'Banana' },
  { key: 'orange', text: 'Orange', disabled: true },
  { key: 'grape', text: 'Grape' },
  { key: 'divider_1', text: '-', itemType: DropdownMenuItemType.Divider },
  { key: 'vegetablesHeader', text: 'Vegetables', itemType: DropdownMenuItemType.Header },
  { key: 'broccoli', text: 'Broccoli' },
  { key: 'carrot', text: 'Carrot' },
  { key: 'lettuce', text: 'Lettuce' },
];
const optionsslob: IDropdownOption[] = [
  { key: 'Global Functions', text: 'Global Functions' },
  { key: 'Downstream', text: 'Downstream' },
  { key: 'Upstream', text: 'Upstream' },
];
const onFormatDate = (date: Date): string => {
  return date.getDate() + '/' + (date.getMonth() + 1) + '/' + (date.getFullYear() % 100);
};
export default class McarForm extends React.Component<IMcarFormProps,IMcarFormState, {}> {

  //checks the url query string and determines the new/edit/display forms
  private detectQueryString(url) 
  {
  var currmode;
  var currid;
  var params = {};
	var parser = document.createElement('a');
	parser.href = url;
	var query = parser.search.substring(1);
	var vars = query.split('&');
	for (var i = 0; i < vars.length; i++) {
		var pair = vars[i].split('=');
		params[pair[0]] = decodeURIComponent(pair[1]);
	}
  Object.keys(params).forEach(key => {
    if(key.toLowerCase()=="id")
    currid=this.getParameterByName(key);
    if(key.toLowerCase()=="mode")
    currmode=this.getParameterByName(key);
  });

  if(currid=="" || currid ==undefined)
  return "new";
  else if(currid!="" || currid!=undefined)
  {
    mcaritemid=currid;    
    return currmode;
  }
  
  }
private getParameterByName(name, url = window.location.href) {
  name = name.replace(/[\[\]]/g, '\\$&');
  var regex = new RegExp('[?&]' + name + '(=([^&#]*)|&|#|$)'),
      results = regex.exec(url);
  if (!results) return null;
  if (!results[2]) return '';
  return decodeURIComponent(results[2].replace(/\+/g, ' '));
}
  constructor(props: IMcarFormProps, state: IMcarFormState) {
    super(props);

    this.state = {
      mode: this.detectQueryString(window.location.href),
      isexistingsolution:false,
      issharepointrelated:false,
      items:[],
      valuesobj:{}      
    };
    this._getisExistingSolution = this._getisExistingSolution.bind(this);   
    this._getO365Systems = this._getO365Systems.bind(this);
    this.createItem = this.createItem.bind(this);
    this.updateItem = this.updateItem.bind(this); 
    this.readItem = this.readItem.bind(this); 
    this._getPeoplePickerItems = this._getPeoplePickerItems.bind(this);  
    this.onSubmit = this.onSubmit.bind(this);
    this.CheckAndAddRequiredField = this.CheckAndAddRequiredField.bind(this);
    this.getParameterByName = this.getParameterByName.bind(this) ;
    this.detectQueryString = this.detectQueryString.bind(this) ;        
  }
  private _getisExistingSolution(ev: React.MouseEvent<HTMLElement>,checked: boolean) {
    if(checked)
    {
      this.setState({isexistingsolution:true});
      newitemobj["Existing_x0020_Solution"]=true;
    }    
    else
    {
      this.setState({isexistingsolution:false});
      newitemobj["Existing_x0020_Solution"]=false;
    }
    
  } 

  private _getO365Systems(ev: React.FormEvent<HTMLElement>,option?: IDropdownOption, index?: number) {
    if(option.key=="SharePoint" && option.selected==true)
    this.setState({issharepointrelated:true})
    else if(option.key=="SharePoint" && option.selected==false)
    this.setState({issharepointrelated:false})     
    if(option.selected==true)
    getSelectedO365Keys.push(option.key); 
    else if(option.selected==false)
    {
     // let seleteditemarr = this.state.seletedprojects;
    let i = getSelectedO365Keys.indexOf(option.key);
    if (i >= 0) {
      getSelectedO365Keys.splice(i, 1);
    }
    }                                 
    newitemobj["With_x0020_which_x0020_O365_x002"]=getSelectedO365Keys;              
  } 

  private _getPeoplePickerItems(items: any[]) {  
    let getSelectedUsers = [];  
    for (let item in items) {  
      getSelectedUsers.push(items[item].id);  
    }       
  }

 //function to add required field div on the controls based on the required fields arr
  private CheckAndAddRequiredField(fieldname)
  {
    let currfieldname=fieldname;
    if(requiredfields.indexOf(currfieldname)>-1)
    {            
        return (<div id={currfieldname} className="form-validation">
               <span>This is a required field</span>
              </div>);
    }
    else
    return null;
  }

//function to fill values in controls depending on mode
private FillValue(controlname,controltype) : any
{  
  if(Object.keys(edititemobj).length>0)
  {
    if(this.state.mode=="edit" || this.state.mode=="display")
    {
      //prefill the multiselect dropdowns
      if(controlname=="With_x0020_which_x0020_O365_x002")
      getSelectedO365Keys=edititemobj[controlname];
      else if(controlname=="Technology_x0020_used")
      getSelectedTechUsedKeys=edititemobj[controlname];
      else if(controlname=="Authentication_x0020_model")
      getSelectedAuthModels=edititemobj[controlname];

      if(controltype=="TextField" || controltype=="DropDown" || controltype=="DropDownMulti")
      return edititemobj[controlname];
      if(controltype=="HyperLink" && Object.keys(edititemobj).length>0)
      return edititemobj[controlname].Url;   
      if(controltype=="PeoplePicker" && Object.keys(edititemobj).length>0)  
      {
        let tempuserarr=[];
        if(edititemobj[controlname]!=null || edititemobj[controlname]!=undefined)
        {
          edititemobj[controlname].forEach(element => 
            {
            tempuserarr.push(element.EMail);
            });
        }        
        return tempuserarr;
      }
      if(controltype=="DateTime" && Object.keys(edititemobj).length>0)
      {
        let tempdate:string;
        tempdate=new Date(edititemobj[controlname]).toString();
        return tempdate;
      }
      if(controltype=="Toggle" && Object.keys(edititemobj).length>0)
      {
        let tempchecked:boolean;
        ischecked=eval(edititemobj[controlname]);
        return ischecked;
      }
    }
  }
}

  private onSubmit() {  
    jQuery(".form-validation").attr("style","display:none !important");
    let validation:boolean=true;
    //check against newitemobj if values are updated
  Object.keys(newitemobj).forEach(key => {    
    if(newitemobj[key]!=edititemobj[key])
    edititemobj[key]=newitemobj[key];
  });
     requiredfields.forEach(field => {

      if(this.state.mode=="new")
      {
        if(field=="If_x0020_yes_x002c__x0020_please" && newitemobj["Does_x0020_your_x0020_solution_x0"]==true )
        {
          if(newitemobj[field]==undefined || newitemobj[field]=="" )
          {
          document.getElementById(field).setAttribute("style","display:block !important");
          validation=false;
          }  
        }
        else if(field=="If_x0020_yes_x002c__x0020_list_x" && newitemobj["Does_x0020_solution_x0020_rely_x"]==true )
        {
          if(newitemobj[field]==undefined || newitemobj[field]=="" )
          {
           document.getElementById(field).setAttribute("style","display:block !important");
           validation=false;
          } 
        }
       else if(field!="If_x0020_yes_x002c__x0020_please" && field!="If_x0020_yes_x002c__x0020_list_x")
       {
        if(newitemobj[field]==undefined || newitemobj[field]=="" )
        {
         document.getElementById(field).setAttribute("style","display:block !important");
         validation=false;
        }
       }
      }
      else if(this.state.mode=="edit")
      {
        if(field=="If_x0020_yes_x002c__x0020_please" && edititemobj["Does_x0020_your_x0020_solution_x0"]==true )
        {
          if(edititemobj[field]==undefined || edititemobj[field]=="" )
          {
          document.getElementById(field).setAttribute("style","display:block !important");
          validation=false;
          }  
        }
        else if(field=="If_x0020_yes_x002c__x0020_list_x" && edititemobj["Does_x0020_solution_x0020_rely_x"]==true )
        {
          if(edititemobj[field]==undefined || edititemobj[field]=="" )
          {
           document.getElementById(field).setAttribute("style","display:block !important");
           validation=false;
          } 
        }
       else if(field!="If_x0020_yes_x002c__x0020_please" && field!="If_x0020_yes_x002c__x0020_list_x")
       {
        if(edititemobj[field]==undefined || edititemobj[field]=="" )
        {
         document.getElementById(field).setAttribute("style","display:block !important");
         validation=false;
        }
       }
      }                      
     });
     if(validation)//passed all validations
     {
       if(this.state.mode=="new")
       this.createItem();
       else if(this.state.mode=="edit")
       this.updateItem();
     }
     
  }


   //function to create the item in Mcar List
   private createItem(): void {  
    this.setState({         
      items: []  
    });  
    
    const body: string = JSON.stringify({  
      'Title': `Item ${new Date()}`,
      'Existing_x0020_Solution':newitemobj["Existing_x0020_Solution"],
      'Release_x0020_Notes':newitemobj["Release_x0020_Notes"],
      'Solution_x0020_Name':newitemobj["Solution_x0020_Name"],
      'Line_x0020_of_x0020_Business':newitemobj["Line_x0020_of_x0020_Business"],
      'Start_x0020_Development_x0020_Ap':newitemobj["Start_x0020_Development_x0020_Ap"],
      'Solution_x0020_architectId':newitemobj["Solution_x0020_architect"],
      'Solution_x0020_SponsorId': newitemobj["Solution_x0020_Sponsor"],
      'Description':newitemobj["Description"],
      'Solution_x0020_OwnerId':newitemobj["Solution_x0020_Owner"],
      'Anticipated_x0020_go_x0020_live_':newitemobj["Anticipated_x0020_go_x0020_live_"],
      'With_x0020_which_x0020_O365_x002':newitemobj["With_x0020_which_x0020_O365_x002"]
    });  
     
    const newbody:string=JSON.stringify(newitemobj);
    this.props.spHttpClient.post(`${this.props.siteUrl}/_api/web/lists/getbytitle('MCARAgile')/items`,  
    SPHttpClient.configurations.v1,  
    {  
      headers: {  
        'Accept': 'application/json;odata=nometadata',  
        'Content-type': 'application/json;odata=nometadata',  
        'odata-version': ''  
      },  
      body: newbody
    })  
    .then((response: SPHttpClientResponse): Promise<IListItem> => {  
      return response.json();  
    })  
    .then((item: IListItem): void => {  
      this.setState({           
        items: []  
      }); 
      alert("Submitted successfully") ;
    }, (error: any): void => {        
      this.setState({           
        items: []  
      }); 
      alert("An error occured"); 
    });  
  }
//function to update item
 private updateItem(): void {  
  this.setState({         
    items: []  
  }); 
  var tempupdateobj={};
  //check against newitemobj if values are updated
  Object.keys(newitemobj).forEach(key => {    
    if(newitemobj[key]!=edititemobj[key])
    tempupdateobj[key]=newitemobj[key];
  });


  const newbody:string=JSON.stringify(tempupdateobj);
  this.props.spHttpClient.post(`${this.props.siteUrl}/_api/web/lists/getbytitle('MCARAgile')/items(${mcaritemid})`,  
  SPHttpClient.configurations.v1,  
  {  
    headers: {  
      'Accept': 'application/json;odata=nometadata',  
      'Content-type': 'application/json;odata=nometadata',  
      'odata-version': '',  
      'IF-MATCH': '*',  
      'X-HTTP-Method': 'MERGE'   
    },  
    body: newbody
  })  
  .then((response: SPHttpClientResponse): Promise<IListItem> => {  
    return response.json();  
  })  
  .then((item: IListItem): void => {  
    this.setState({           
      items: []  
    }); 
   // alert("Updated successfully") ;
  }, (error: any): void => {        
    this.setState({           
      items: []  
    }); 
    alert("Updated successfully") ; 
  });  
}
  //function to fetch list item using id
  private readItem(): void {  
    this.setState({        
      items: []  
    });  
       
    this.props.spHttpClient.get(`${this.props.siteUrl}/_api/web/lists/getbytitle('MCARAgile')/items(${mcaritemid})?$expand=Solution_x0020_Owner,Business_x0020_IRM_x0020_represe,Design_x0020_author,Solution_x0020_architect,Solution_x0020_Sponsor,Support_x0020_contact_x0020_deta&$select=*,Solution_x0020_Owner/EMail,Business_x0020_IRM_x0020_represe/EMail,Design_x0020_author/EMail,Solution_x0020_architect/EMail,Solution_x0020_Sponsor/EMail,Support_x0020_contact_x0020_deta/EMail`,  
    SPHttpClient.configurations.v1,  
          {  
            headers: {  
              'Accept': 'application/json;odata=nometadata',  
              'odata-version': ''  
            }  
          })      
      .then((response: SPHttpClientResponse): Promise<IListItem> => {  
        return response.json();  
      })  
      .then((item: IListItem): void => {  
        edititemobj=item;        
        console.log("edititemobj : "+edititemobj);
        this.setState({             
          items: [],
          valuesobj:item
        });  
      }, (error: any): void => {  
        this.setState({             
          items: []  
        });  
      });  
  }
  componentWillMount() {

    if(this.state.mode=="edit" || this.state.mode=="display")
    this.readItem();
 }
  public render(): React.ReactElement<IMcarFormProps> {
    return (
      <div className={ styles.mcarForm }>
        <div className={ styles.container }>
        <form onSubmit={this.onSubmit}>
        <div className={ styles.row } >
          <div className={ styles.column }>
          <Image src="" alt="Example with no image fit value and no height or width is specified."/>
          </div>
          <div className={ styles.column }>
          <Label style={{ fontSize: "16pt" }} className={styles.labelHeader}>Microsoft O365 Customisations and Application Registry</Label>
          </div>
        </div>
          <div className={ styles.row } >
            <div className={ styles.column }>                        
              <Toggle 
              label="Existing Solution" 
              onChange={this._getisExistingSolution}                            
              onText="Yes" 
              offText="No" 
              /> 
              { this.state.isexistingsolution?                     
              [<Dropdown                              
                placeholder="Select an option"
                label="Select SPCAR ID"              
                options={optionsspcarid}                
              />,  
              <TextField 
                 name="ReleaseNotesTxtField"
                 label="Release Notes"                  
                 multiline rows={3} 
                 onChange={(e,val) => newitemobj["Release_x0020_Notes"]=val}              
              />]
              :null  
              }                        
              <Toggle 
              label="Start Development Approval" 
              onText="Yes" 
              offText="No" 
              onChange={(e,checked)=>{newitemobj["Start_x0020_Development_x0020_Ap"]=checked}}
              />
              <TextField label="Solution Status" disabled defaultValue="Draft" />              
            </div>
          </div>
         
          <div className={ styles.row } >
            <div className={ styles.column }>            
            <Label className={styles.labelHeader}>General Solution Fields</Label>
            <TextField 
            label="Solution Name " 
            required 
            onChange={(e,val) => newitemobj["Solution_x0020_Name"]=val} 
            defaultValue={this.FillValue("Solution_x0020_Name","TextField")}             
            />
            {this.CheckAndAddRequiredField("Solution_x0020_Name")}
            <Dropdown
                placeholder="Select an option"
                label="Line Of Business"
                options={optionsslob}    
                onChange={(e,val) => newitemobj["Line_x0020_of_x0020_Business"]=val.key}
                required  
                defaultSelectedKey= {this.FillValue("Line_x0020_of_x0020_Business","DropDown")}                           
              />
              {this.CheckAndAddRequiredField("Line_x0020_of_x0020_Business")}
              <PeoplePicker
                context={this.props.context}
                titleText="Solution Owner"
                personSelectionLimit={10}
                placeholder="Enter names.."
                groupName={""} // Leave this blank in case you want to filter from all users
                showtooltip={true}   
                required={true}             
                disabled={false}
                ensureUser={true}    
                defaultSelectedUsers={this.FillValue("Solution_x0020_Owner","PeoplePicker")}                        
                onChange={(items) =>{
                  let getSelectedUsers = [];  
                  for (let item in items) {  
                  getSelectedUsers.push(items[item].id);  
                  newitemobj["Solution_x0020_OwnerId"]=getSelectedUsers;
                  } 
                 }                   
                 }
                showHiddenInUI={false}
                principalTypes={[PrincipalType.User]}
                resolveDelay={1000} />
               {this.CheckAndAddRequiredField("Solution_x0020_OwnerId")}

                <PeoplePicker
                context={this.props.context}
                titleText="Solution Sponsor"
                personSelectionLimit={10}
                placeholder="Enter names.."
                groupName={""} // Leave this blank in case you want to filter from all users
                showtooltip={true}   
                ensureUser={true}  
                required={true}             
                disabled={false}
                defaultSelectedUsers={this.FillValue("Solution_x0020_Sponsor","PeoplePicker")} 
                //onChange={this._getPeoplePickerItems}
                onChange={(items) =>{
                  let getSelectedUsers = [];  
                  for (let item in items) {  
                  getSelectedUsers.push(items[item].id);  
                  newitemobj["Solution_x0020_SponsorId"]=getSelectedUsers;
                  } 
                 }                   
                 }
                showHiddenInUI={false}
                principalTypes={[PrincipalType.User]}
                resolveDelay={1000} />
                {this.CheckAndAddRequiredField("Solution_x0020_SponsorId")}
                <TextField 
                label="Description (Including Business Case)" 
                onChange={(e,val) => newitemobj["Description"]=val}
                required 
                multiline 
                defaultValue={this.FillValue("Description","TextField")} 
                rows={3} />
                 {this.CheckAndAddRequiredField("Description")}
                <TextField 
                label="Scope & Approach" 
                onChange={(e,val) => newitemobj["Business_x0020_case"]=val}
                required 
                defaultValue={this.FillValue("Business_x0020_case","TextField")} 
                multiline 
                rows={3} 
                />
                {this.CheckAndAddRequiredField("Business_x0020_case")}
                <PeoplePicker
                context={this.props.context}
                titleText="Solution Architect"
                personSelectionLimit={10}
                placeholder="Enter names.."
                groupName={""} // Leave this blank in case you want to filter from all users
                showtooltip={true}   
                ensureUser={true}  
                required={false}             
                disabled={false}
                defaultSelectedUsers={this.FillValue("Solution_x0020_architect","PeoplePicker")} 
               onChange={(items) =>{
                let getSelectedUsers = [];  
                for (let item in items) {  
                getSelectedUsers.push(items[item].id);  
                newitemobj["Solution_x0020_architectId"]=getSelectedUsers;
                } 
               }                   
               }
                showHiddenInUI={false}
                principalTypes={[PrincipalType.User]}
                resolveDelay={1000} />
                 <PeoplePicker
                context={this.props.context}
                titleText="Design Author"
                personSelectionLimit={10}
                placeholder="Enter names.."
                groupName={""} // Leave this blank in case you want to filter from all users
                showtooltip={true}   
                required={false}    
                ensureUser={true}           
                disabled={false}
                defaultSelectedUsers={this.FillValue("Design_x0020_author","PeoplePicker")} 
                onChange={(items) =>{
                  let getSelectedUsers = [];  
                  for (let item in items) {  
                  getSelectedUsers.push(items[item].id);  
                  newitemobj["Design_x0020_authorId"]=getSelectedUsers;
                  } 
                 }                   
                 }
                showHiddenInUI={false}
                principalTypes={[PrincipalType.User]}
                resolveDelay={1000} />
                <TextField 
                label="Technical Design (Insert Link)" 
                placeholder="https://.." 
                defaultValue={this.FillValue("Technical_x0020_design","HyperLink")} 
                onChange={(e,val) => 
                  {
                    var link = new Object();
                    link["Description"]="";
                    link["Url"]=val;
                    newitemobj["Technical_x0020_design"]=link;
                  }                                  
                }
                />
                <TextField 
                label="Support" 
                multiline 
                onChange={(e,val) => newitemobj["Support"]=val}
                defaultValue={this.FillValue("Support","TextField")} 
                rows={3} 
                />
                <PeoplePicker
                context={this.props.context}
                titleText="Support Contact(s) Detail"
                personSelectionLimit={10}
                placeholder="Enter names.."
                groupName={""} // Leave this blank in case you want to filter from all users
                showtooltip={true}   
                required={false}             
                disabled={false}
                ensureUser={true} 
                defaultSelectedUsers={this.FillValue("Support_x0020_contact_x0020_deta","PeoplePicker")} 
                onChange={(items) =>{
                  let getSelectedUsers = [];  
                  for (let item in items) {  
                  getSelectedUsers.push(items[item].id);  
                  newitemobj["Support_x0020_contact_x0020_detaId"]=getSelectedUsers;
                  } 
                 }                   
                 }
                showHiddenInUI={false}
                principalTypes={[PrincipalType.User]}
                resolveDelay={1000} />
                <TextField 
                label="APEX ID" 
                onChange={(e,val) => newitemobj["APEX_x0020_ID"]=val}
                defaultValue={this.FillValue("APEX_x0020_ID","TextField")}
                placeholder="Enter APEX Id" 
                />
                <TextField 
                label="APEX Link"    
                onChange={(e,val) => newitemobj["APEX_x0020_Link"]=val}   
                defaultValue={this.FillValue("APEX_x0020_Link","TextField")}                         
                placeholder="https://.." 
                />
                <PeoplePicker
                context={this.props.context}
                titleText="Business IRM Representative for solution"
                personSelectionLimit={10}
                placeholder="Enter names.."
                groupName={""} // Leave this blank in case you want to filter from all users
                showtooltip={true}   
                required={false}      
                ensureUser={true}        
                disabled={false}
                defaultSelectedUsers={this.FillValue("Business_x0020_IRM_x0020_represe","PeoplePicker")} 
                onChange={(items) =>{
                  let getSelectedUsers = [];  
                  for (let item in items) {  
                  getSelectedUsers.push(items[item].id);  
                  newitemobj["Business_x0020_IRM_x0020_represeId"]=getSelectedUsers;
                  } 
                 }                   
                 }
                showHiddenInUI={false}
                principalTypes={[PrincipalType.User]}
                resolveDelay={1000} />
                <TextField 
                label="Provide link to IRM approval (IRM Approval is mandatory for applications that require APEX registration)" 
                defaultValue={this.FillValue("Upload_x0020_IRM_x0020_approval_","HyperLink")}                
                onChange={(e,val) => 
                  {
                    var link = new Object();
                    link["Description"]="";
                    link["Url"]=val;
                    newitemobj["Upload_x0020_IRM_x0020_approval_"]=link;
                  }                                  
                }
                placeholder="https://.." 
                />
                <DatePicker
                label="Anticipated Go Live Date"
                isMonthPickerVisible={true}
                showGoToToday={false}                
                formatDate={onFormatDate}
                //value={new Date('2020-11-27T18:30:00Z')}
                //defaultValue={new Date('2020-11-27T18:30:00Z')} 
                value={new Date(this.FillValue("Anticipated_x0020_go_x0020_live_","DateTime"))} 
                onSelectDate={(dateitem) =>{              
                  newitemobj["Anticipated_x0020_go_x0020_live_"]=dateitem;
                  }                                    
                 }
              //  allowTextInput={true}
              //  value={this.state.currselecteddate}
              />
              <DatePicker
                label="Anticipated Retirement Date"
                isMonthPickerVisible={true}
                showGoToToday={false}                
                formatDate={onFormatDate}
                value={new Date(this.FillValue("Anticipated_x0020_retirement_x00","DateTime"))} 
                onSelectDate={(dateitem) =>{              
                  newitemobj["Anticipated_x0020_retirement_x00"]=dateitem;
                  }                                    
                 }
              //  allowTextInput={true}
              //  value={this.state.currselecteddate}
              />
              <Dropdown
                placeholder="Select an option"
                label="Is this an enterprise wide solution or scoped to a sub-set of user"
                onChange={(e,val) =>                   
                    newitemobj["Is_x0020_this_x0020_an_x0020_ent"]=val.key                                                     
                }
                defaultSelectedKey={this.FillValue("Is_x0020_this_x0020_an_x0020_ent","DropDown")}               
                options={[
                  { key: 'Enterprise Wide Solution', text: 'Enterprise Wide Solution'},
                  { key: 'Scoped to a sub-set of user', text: 'Scoped to a sub-set of user' },                
                ]}                                              
              />
              <Dropdown
                placeholder="Select an option"
                label="With which O365 System(s) will you interface?"
                onChange={this._getO365Systems}
                required
                defaultSelectedKeys={this.FillValue("With_x0020_which_x0020_O365_x002","DropDownMulti")}
                multiSelect
                options={[
                  { key: 'Exchange', text: 'Exchange'},
                  { key: 'SharePoint', text: 'SharePoint' },                
                  { key: 'OneDrive', text: 'OneDrive' },  
                  { key: 'Yammer', text: 'Yammer' },  
                ]}                              
              />
                {this.CheckAndAddRequiredField("With_x0020_which_x0020_O365_x002")}
            </div>
          </div>    

          <div className={ styles.row } >
            <div className={ styles.column }>  
            <Accordion allowZeroExpanded={true}>
            <AccordionItem>
                <AccordionItemHeading>  
                <AccordionItemButton>  
                  <Label className={styles.labelHeader}>Initial Scoping Questions</Label>
                </AccordionItemButton>
              </AccordionItemHeading>
              <AccordionItemPanel>
              <Dropdown
                placeholder="Select an option"
                label="What classification does the data consumed by the application have?"  
                onChange={(e,val) => newitemobj["What_x0020_Classification_x0020_"]=val.key}    
                defaultSelectedKey={this.FillValue("What_x0020_Classification_x0020_","DropDown")}          
                options={[
                  { key: 'Unrestricted', text: 'Unrestricted'},
                  { key: 'Restricted', text: 'Restricted' },
                  { key: 'Confidential', text: 'Confidential' },                    
                ]}                              
              />
              <Dropdown
                placeholder="Select an option"
                label="What kind of operation does this application do to data?"
                onChange={(e,val) => newitemobj["What_x0020_kind_x0020_of_x0020_O"]=val.text}  
                defaultSelectedKey={this.FillValue("What_x0020_kind_x0020_of_x0020_O","DropDown")}             
                options={[
                  { key: 'Create, read, update, and delete (In-Place Data Operation)', text: 'Create, read, update, and delete (In-Place Data Operation)'},
                  { key: 'Direct-Move (Pass Through)', text: 'Direct-Move (Pass Through)' },                
                  { key: 'Collect/Process/Publish (ETL)  (new data will be added or existing data will be modified and moved)', text: 'Collect/Process/Publish (ETL)  (new data will be added or existing data will be modified and moved)' },  
                  { key: 'Business Process/Approval Workflow', text: 'Business Process/Approval Workflow' },  
                ]}                              
              />
              <Dropdown
                placeholder="Select an option"
                label="Data move (between containers of data with different classification)"
                onChange={(e,val) => newitemobj["Data_x0020_Move_x0020__x0028_Bet"]=val.text}
                defaultSelectedKey={this.FillValue("Data_x0020_Move_x0020__x0028_Bet","DropDown")}  
                options={[
                  { key: 'Data Does not Move', text: 'Data Does not Move'},
                  { key: 'Move/Publish From Unrestricted to Restricted/Confidential', text: 'Move/Publish From Unrestricted to Restricted/Confidential' },                
                  { key: 'Move/Publish From Restricted to Restricted/Confidential', text: 'Move/Publish From Restricted to Restricted/Confidential' },  
                  { key: 'Move/Publish From Confidential to Confidential', text: 'Move/Publish From Confidential to Confidential' },  
                  { key: 'Move/Publish From Confidential to Restricted', text: 'Move/Publish From Confidential to Restricted' },  
                ]}                                              
              />
               <Dropdown
                placeholder="Select options"
                label="Technology(s) used"
                required                
                multiSelect
                defaultSelectedKeys={this.FillValue("Technology_x0020_used","DropDownMulti")} 
                onChange={(e,item) => 
                  {
                  if(item.selected==true)
                  getSelectedTechUsedKeys.push(item.text); 
                  else if(item.selected==false)
                  {                  
                  let i = getSelectedTechUsedKeys.indexOf(item.text);
                  if (i >= 0) {
                    getSelectedTechUsedKeys.splice(i, 1);
                  }
                  }                   
                  newitemobj["Technology_x0020_used"]=getSelectedTechUsedKeys;
                  }
                }
                options={[
                  { key: 'Web application', text: 'Web application'},
                  { key: 'Native application (script, tool, mobile app)', text: 'Native application (script, tool, mobile app)' },                
                  { key: 'Office Add ins', text: 'Office Add ins' },  
                  { key: 'SharePoint Add ins', text: 'SharePoint Add ins' },  
                  { key: 'SPFx (webpart; application customiser, field customiser, cmd sets)', text: 'SPFx (webpart; application customiser, field customiser, cmd sets)' },  
                  { key: 'Infopath', text: 'Infopath' },  
                  { key: 'Workflow', text: 'Workflow' },  
                  { key: 'PowerApps', text: 'PowerApps' },  
                  { key: 'MS Flow', text: 'MS Flow' },  
                ]}                
              />
               {this.CheckAndAddRequiredField("Technology_x0020_used")}
              <Dropdown
                placeholder="Select options"
                label="Solution Type"
                required      
                onChange={(e,val) => newitemobj["Solution_x0020_type"]=val.text} 
                defaultSelectedKey={this.FillValue("Solution_x0020_type","DropDown")}                          
                options={[
                  { key: 'Application built using O365 suite (e.g. Workflow, SPFx, PowerApp, Flow (standard O365 connectors))', text: 'Application built using O365 suite (e.g. Workflow, SPFx, PowerApp, Flow (standard O365 connectors))'},
                  { key: 'Application outside O365 suite that is connecting to O365 (3rd party connector or add-in, Azure AD registered application, Flow (Premium, Custom, Azure & non-Microsoft connectors))', text: 'Application outside O365 suite that is connecting to O365 (3rd party connector or add-in, Azure AD registered application, Flow (Premium, Custom, Azure & non-Microsoft connectors))' },                                   
                ]}                
              />
              {this.CheckAndAddRequiredField("Solution_x0020_type")}
              </AccordionItemPanel>
              </AccordionItem>
            </Accordion>
            </div>
          </div>
          
          <div className={ styles.row } >
            <div className={ styles.column }>     
            {
              this.state.issharepointrelated?
             [ <Label className={styles.labelHeader}>SharePoint</Label>,
              <TextField 
              label="URLs Development"               
              onChange={(e,val) => 
                {
                  var link = new Object();
                  link["Description"]="";
                  link["Url"]=val;
                  newitemobj["URLs_x0020_Development"]=link;
                }                                  
              }
              placeholder="https://.."
              />,
              <TextField label="URLs Acceptance" placeholder="https://.."               
              onChange={(e,val) => 
                {
                  var link = new Object();
                  link["Description"]="";
                  link["Url"]=val;
                  newitemobj["URLs_x0020_Acceptance"]=link;
                }                                  
              }
              />,
              <TextField label="URLs Production" placeholder="https://.."               
              onChange={(e,val) => 
                {
                  var link = new Object();
                  link["Description"]="";
                  link["Url"]=val;
                  newitemobj["URLs_x0020_Production"]=link;
                }                                  
              }
              />,
              <Toggle label="Does your solution use any other SharePoint customisations other than those listed in the 'technology used' field?" onText="Yes" offText="No" onChange={(e,val) => newitemobj["Does_x0020_your_x0020_solution_x0"]=val} />,
              <TextField label="If yes, please specify" required multiline rows={3}  onChange={(e,val) => newitemobj["If_x0020_yes_x002c__x0020_please"]=val}
              />,
              <div id="If_x0020_yes_x002c__x0020_please" className="form-validation">
               <span>This is a required field</span>
              </div>,
              <TextField label="Please add link to evidence that the SharePoint Functional Site Owner has approved for your application to be deployed on their site collection (Insert Link)" placeholder="https://.."               
              onChange={(e,val) => 
                {
                  var link = new Object();
                  link["Description"]="";
                  link["Url"]=val;
                  newitemobj["Please_x0020_attach_x0020_eviden"]=link;
                }                                  
              }
              />]
              :null
            } 
                  
              
            </div>
          </div>
          <div className={ styles.row } >
            <div className={ styles.column }>            
              <Label className={styles.labelHeader}>Search</Label>              
              <Toggle 
              label="Does your solution require any change to Search configuration?" 
              defaultChecked={this.FillValue("Does_x0020_your_x0020_solution_x","Toggle")}                  
              onText="Yes" 
              offText="No" 
              onChange={(e,val) => newitemobj["Does_x0020_your_x0020_solution_x"]=val}                         
              />                            
            </div>
          </div>
          <div className={ styles.row } >
            <div className={ styles.column }>            
              <Label className={styles.labelHeader}>Integration - non-Powerapps/Flow</Label>
              <TextField 
              label="Name of the system connecting to O365" 
              defaultValue={this.FillValue("Name_x0020_of_x0020_the_x0020_sy","TextField")} 
              placeholder="" onChange={(e,val) => newitemobj["Name_x0020_of_x0020_the_x0020_sy"]=val}
              />     
              <Dropdown
                placeholder="Select an option"
                label="Authentication Model(s)"
                multiSelect     
                defaultSelectedKeys={this.FillValue("Authentication_x0020_model","DropDownMulti")}            
                onChange={(e,item) => 
                  {
                  if(item.selected==true)
                  getSelectedAuthModels.push(item.text); 
                  else if(item.selected==false)
                  {                  
                  let i = getSelectedAuthModels.indexOf(item.text);
                  if (i >= 0) {
                    getSelectedAuthModels.splice(i, 1);
                  }
                  }                   
                  newitemobj["Authentication_x0020_model"]=getSelectedAuthModels;
                  }
                }
                options={[
                  { key: 'Azure AD Application', text: 'Azure AD Application'},
                  { key: 'ACS Client Id/Secret', text: 'ACS Client Id/Secret' },                
                  { key: 'Cloud identity', text: 'Cloud identity' },                    
                ]}                              
              />       
              <TextField 
              label="Please ensure that the design attached to this form includes details of the permissions you require" 
              defaultValue={this.FillValue("Please_x0020_ensure_x0020_that_x","TextField")}   
              multiline rows={3} 
              onChange={(e,val) => newitemobj["Please_x0020_ensure_x0020_that_x"]=val} 
              />  
              <Toggle 
              label="Does solution rely on or call external services?" 
              defaultChecked={this.FillValue("Does_x0020_solution_x0020_rely_x","Toggle")}               
              onText="Yes" 
              offText="No" 
              onChange={(e,val) => newitemobj["Does_x0020_solution_x0020_rely_x"]=val}
              /> 
              <TextField 
              label="If yes, list them" 
              required 
              multiline 
              defaultValue={this.FillValue("If_x0020_yes_x002c__x0020_list_x","TextField")}  
              rows={3} 
              onChange={(e,val) => newitemobj["If_x0020_yes_x002c__x0020_list_x"]=val} 
              />  
              {this.CheckAndAddRequiredField("If_x0020_yes_x002c__x0020_list_x")}  
              <Toggle 
              label="Does solution rely on access to Office Graph or Microsoft Graph APIs?" 
              defaultChecked={this.FillValue("Does_x0020_solution_x0020_rely_x0","Toggle")}               
              onText="Yes" 
              offText="No" 
              onChange={(e,val) => newitemobj["Does_x0020_solution_x0020_rely_x0"]=val}/>                                    
            </div>
          </div>
          <div className={ styles.row } >
            <div className={ styles.column }>  
            <PrimaryButton 
            text="Submit" 
            onClick={this.onSubmit} 
            allowDisabledFocus 
            //disabled={disabled} 
            //checked={checked} 
            />   
            
            <DefaultButton 
            text="Cancel" 
            //onClick={_alertClicked} 
            allowDisabledFocus 
            //disabled={disabled} 
            //checked={checked} 
            />   
            </div>                                                    
            
          </div>
        </form>
        </div>      
      </div>
    );
  }
  
}
