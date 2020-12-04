/////////////////////////////////////////////////////////////////////////
//  Title: HelloWorld WebPart
//  Author: Lucas Johnson
//
//  Description: For use on HVAF SharePoint Site. Provides interface
//  and backend for managing GSR cases. 
//
//  Version: 1.1 
//  Changelog: Fixed refresh page breaking submit button bug, refactored
//  entire codebase, aesthetic overhaul.
/////////////////////////////////////////////////////////////////////////


/////////////////////////////////////////////////////////////////////////
//  Section 1 - Imports, Exports, & Initialization
/////////////////////////////////////////////////////////////////////////

//  Import version control, required for dataVersion() getter
import { Version } from '@microsoft/sp-core-library';

//  Import config elements for config box, used to switch between Front Desk and Volunteer Mode
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';

//  Import WebPart elements, powers WebPart
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape, extend } from '@microsoft/sp-lodash-subset';

//  Import styles
import styles from './HelloWorldWebPart.module.scss';
import * as strings from 'HelloWorldWebPartStrings';

//  Import SPHttlClient and sp pnp for making API calls
import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';
import {
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';
import { sp } from "@pnp/sp/presets/all";


//  Export modeToggle bool, used to store the mode
export interface IHelloWorldWebPartProps {
  modeToggle: boolean;
}

//  Export list elements
export interface ISPLists {
  value: ISPList[];
}
export interface ISPList {
  Title: string;
  Id: string;
}

//  Initialize maxFoodPoints, declared here as a global variable so that it can be altered from anywhere
var maxFoodPoints = 100;


//  Definition of the entire WebPart
export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {

/////////////////////////////////////////////////////////////////////////
//  Section 2: Frontend
/////////////////////////////////////////////////////////////////////////

  public render(): void {
    
    if (this.properties.modeToggle) { //  True = Front Desk Mode
      this.domElement.innerHTML = `
      <div class="${ styles.helloWorld }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <style>
                .modal {
                  display: none; 
                  position: fixed; 
                  z-index: 1; 
                  padding-top: 100px; 
                  left: 0;
                  top: 0;
                  width: 100%; 
                  height: 100%; 
                  overflow: auto;
                  background-color: rgb(0,0,0);
                  background-color: rgba(0,0,0,0.4); 
                }
                
                .modal-content {
                  background-color: #fefefe;
                  margin-top: 56px;
                  margin-left: 250px;
                  font-size: 150%;
                  padding: 20px;
                  border: 1px solid #888;
                  width: 25%;
                }                            
              </style>


              <div id="addRefresh">
                
              </div>

              <div id="myModal" class="modal">

              <div class="modal-content">
                <p style="color:black;">You have just assigned a veteran to a volunteer.</p>
                <p style="color:black;">Press refresh below to continue:</p>
                <button style="font-size: 30px;" onClick="window.location.reload();">Refresh Page</button>
  
              </div>


              </div>
              <p class="${ styles.description }">Loading from ${escape(this.context.pageContext.web.title)}, using Front Desk Mode</p>
              <p>Page URL: ${escape(this.context.pageContext.web.absoluteUrl)} </p>
              <table>
                <tr>
                  <td>Enter Item ID:</td>
                  <td><input type="text" id="txtItemID" size="20" /></td>
                  <td> </td>
                  <td><input type="button" id="btnClickCheck" value="Check Record" /></td>
                </tr>
              </table>
              <br>
              <textarea id="vetRecordTextFeild" rows="8" cols="47" style="resize: none" readonly placeholder='Veteran Info will appear here. Please check to make sure you have selected the right person, then hit "Assign to Volunteers" below.'></textarea>

              <p>This will assign the veteran to the Volunteers.</p>
              <input type="button" id="btnClickSubmit" value="Assign to Volunteers" style="width: 80%"/>

              
            </div>
          </div>
        </div>
      </div>`;
    
    console.log("Front Desk Mode");
    this._captureClickFrontDesk();  //  Adds event listeners to all buttons so that user action on the frontend can be passed to the backend logic
    }


    if (!this.properties.modeToggle) { // False = Volunteer Mode
          this.domElement.innerHTML = `
          <div class="${ styles.helloWorld}">
            <div class="${styles.container}">
              <div class="${styles.row}">
                <div class="${styles.column}">

                <style>
                  .modal {
                    display: none; 
                    position: fixed; 
                    z-index: 1; 
                    padding-top: 100px; 
                    left: 0;
                    top: 0;
                    width: 100%; 
                    height: 100%; 
                    overflow: auto;
                    background-color: rgb(0,0,0);
                    background-color: rgba(0,0,0,0.4); 
                  }
                  
                  
                  .modal-content {
                    background-color: #fefefe;
                    margin: auto;
                    font-size: 300%;
                    padding: 20px;
                    border: 1px solid #888;
                    width: 75%;
                  } 
                  
                  th {
                    border-color: black;
                    border-style: solid;
                    border-width: 1px;
                    font-family: Arial, sans-serif;
                    font-weight: normal;
                    overflow: hidden;
                    padding: 10px 5px;
                    word-break: normal;
                  }
                    
                  .table-header-1 {
                    font-size: 24px;
                    text-align: left;
                    vertical-align: top;
                  }
                    
                  .table-header-2 {
                    font-size: 54px;
                    text-align: center;
                    vertical-align: center;
                  }
                    
                  .table-header-3 {
                    font-size: 14px;
                    text-align: left;
                    vertical-align: top;
                  }
                    
                  .table-header-4 {
                    font-size: 24px;
                    text-align: left;
                    vertical-align: top;
                  }
                    
                  .block {
                    width: 50px;
                    height: 50px;
                    border: none;
                    background-color: black;
                    color: white;
                    font-size: 30px;
                    cursor: pointer;
                    text-align: center;
                  }
                </style>

                <div id="addRefresh2">                

                </div>

                <div id="myModal2" class="modal">

                  <div class="modal-content">
                    <p style="color:black;">You have just updated a GSR record.</p>
                    <p style="color:black;">Press refresh below to continue:</p>
                    <button style="font-size: 50px;" onClick="window.location.reload();">Refresh Page</button>        
                  </div>

                 </div>

                <p class="${styles.description}">Loading from ${escape(this.context.pageContext.web.title)}, using Volunteer Mode</p>
                <p>Page URL: ${escape(this.context.pageContext.web.absoluteUrl)} </p>
                <table>
                    <tr>
                        <td>Enter Item ID:</td>
                        <td><input type="text" id="txtItemID2" size="20" inputmode="decimal" style="font-size: 24px;width: 195px;height: 44px;"/></td>
                        <td> </td>
                        <td><input type="button" id="btnClickCheck2" value="Check Record" style="width: 97px;height: 50px;" /></td>
                    </tr>
                </table>
                <br>
                <textarea id="vetRecordTextFeild2" rows="8" cols="63" style="resize: none" readonly placeholder='Veteran Info will appear here. Please enter the required values, then hit "Submit Record" below.'></textarea>
  
                <!--I AM GENUINELY SORRY ABOUT THE TABLE BELOW-->
                <!--You can't add scripts in here so I had to do everything for the buttons in-line-->

                <p></p>
                <table style="table-layout: fixed;width: 400px;border-collapse: collapse;border-spacing: 0;">
                <colgroup>
                    <col style="width: 50%">
                    </col>
                    <col style="width: 34%">
                    </col>
                    <col style="width: 16%">
                    </col>
                </colgroup>
                <thead>
                    <tr> 
                        <th class="table-header-1" id="Row1Text">Food Points</th>
                        <th class="table-header-2" id="Row1Val">0</th>
                        <th class="table-header-3">
                          <button class="block" onMouseOver="this.style.backgroundColor='grey';this.style.color='black'" onMouseOut="this.style.backgroundColor='black';this.style.color='white'" onclick='document.getElementById("Row1Val").innerHTML = parseInt(document.getElementById("Row1Val").innerHTML)+1;if(parseInt(document.getElementById("Row1Val").innerHTML) > ${JSON.stringify(maxFoodPoints)} ){ document.getElementById("Row1Val").innerHTML = ${JSON.stringify(maxFoodPoints)} };'>▲</button>

                            <p></p>

                          <button class="block" onMouseOver="this.style.backgroundColor='grey';this.style.color='black'" onMouseOut="this.style.backgroundColor='black';this.style.color='white'" onclick='document.getElementById("Row1Val").innerHTML = parseInt(document.getElementById("Row1Val").innerHTML)-1;if(parseInt(document.getElementById("Row1Val").innerHTML) < 0){ document.getElementById("Row1Val").innerHTML = 0};'>▼</button>
                        </th>
                    </tr>

                    <tr>
                        <th class="table-header-1" id="Row2Text">Hygiene Points</th>
                        <th class="table-header-2" id="Row2Val">0</th>
                        <th class="table-header-3">
                          <button class="block" onMouseOver="this.style.backgroundColor='grey';this.style.color='black'" onMouseOut="this.style.backgroundColor='black';this.style.color='white'" onclick='document.getElementById("Row2Val").innerHTML = parseInt(document.getElementById("Row2Val").innerHTML)+1;if(parseInt(document.getElementById("Row2Val").innerHTML) > 10){ document.getElementById("Row2Val").innerHTML = 10};'>▲</button>

                            <p></p>

                          <button class="block" onMouseOver="this.style.backgroundColor='grey';this.style.color='black'" onMouseOut="this.style.backgroundColor='black';this.style.color='white'" onclick='document.getElementById("Row2Val").innerHTML = parseInt(document.getElementById("Row2Val").innerHTML)-1;if(parseInt(document.getElementById("Row2Val").innerHTML) < 0){ document.getElementById("Row2Val").innerHTML = 0};'>▼</button>
                        </th>
                    </tr>

                    <tr>
                        <th class="table-header-4" ><p>Lbs. of Food</p><input type="text" id="lbsFoodVal" size="20" inputmode="decimal" style="font-size: 24px;width: 95%;height: 44px;"/></th>
                        <th colspan="2" class="table-header-4"><p>Lbs. of Clothes</p><input type="text" id="lbsClothesVal" size="20" inputmode="decimal" style="font-size: 24px;width: 95%;height: 44px;"/></th>
    
                    </tr>

                </thead>
            </table>

                <p>This will assign the above values to the Record and mark it Closed.</p>
                <input type="button" id="btnClickSubmit2" value="Submit Record" style="width: 400px;height: 50px"/>

            </div>
          </div>
        </div>
      </div>
      `;

      console.log("Volunteer Mode");

      this._captureClickVolunteer();  //  Adds event listeners to all buttons so that user action on the frontend can be passed to the backend logic
    }
    
  }

  //  Project version getter
  protected get dataVersion(): Version {
    return Version.parse('1.1');
  }

  //  Defines config options
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
  return {
    pages: [
      {
        header: {
          description: 'Custom Web Part used to move & update list items on the HVAF WebForm Responses List. Front Desk Mode allows a front desk employee to assign a Vet to the Volunteers. Volunteer Mode allows a volunteer to enter visit details for a Vet.'
        },
        groups: [
          {
            groupName: strings.BasicGroupName,
            groupFields: [
            PropertyPaneToggle('modeToggle', {
              label: 'Mode',
              onText: 'Front Desk Mode',
              offText: 'Volunteer Mode'
            })
          ]
          }
        ]
      }
    ]
  };
  }


/////////////////////////////////////////////////////////////////////////
//  Section 3: Backend Logic
/////////////////////////////////////////////////////////////////////////

  //  Front Desk Logic
  private _captureClickFrontDesk(): void{
    
    //  Sets up an asynchonous function which inserts a refresh button after 300,000ms, (5min)
    setTimeout(() => { document.getElementById("addRefresh").innerHTML = '<div>Some information may be out of date. Click here to refresh! <button onClick="window.location.reload();">Refresh Page</button></div>'; }, 300000);

    //  Add a click event listener to the Check Record Button, calling a function which handles the input upon the click
    document.querySelector("#btnClickCheck").addEventListener("click", handleClickCheck);
    
    document.querySelector("#btnClickSubmit").addEventListener("click", (e:Event) => {
      if (isIDValid()) {
        //  Updates the Item using the TrueID
        this._updateItemAssign(vetTrueID);
      }
    });

    //  vetTitleID is the "ID - Real" of a particular entry. SharePoint uses a globally unique "Id" for API calls though, so we must differentiate (and eventually translate) between the two
    //  We set each to -1 to preform validity checks later on. If translation fails, for example, vetTrueID will remain negative, and will throw an error
    var vetTitleID = -1;
    var vetTrueID = -1;
  

    function handleClickCheck() {
      //  TypeScript requires we cast certain elements as "HTML" elements, so we set up a temporary landing zone for one such value (the input for the "Check Record" button, i.e. the "ID - Real")
      var vetTitleIDTEMP : number = parseFloat((<HTMLInputElement>document.getElementById("txtItemID")).value);
      vetTitleID = vetTitleIDTEMP;
      
      //  In case the input is blank (i.e. the user has not yet entered a value), we need to throw an error. If the input is blank, the above will have thrown an error, so we must now 
      //  grab the value as a string
      var vetTitleIDSTRING = (<HTMLInputElement>document.getElementById("txtItemID")).value; 
      
      //  If the string is empty, we throw an error
      if (vetTitleIDSTRING == '') {
        document.querySelector("#vetRecordTextFeild").innerHTML = `Error: You must enter an Item ID`;
        vetTitleID = -1;
        return false;
      }

      //  If we've gotten this far, we can now begin the translation between the "ID - Real" aka vetTitleID, and the "Id" aka vetTrueID
      //  We do this by making an API call filtering the items by title, and search for entries that match our vetTitleID.
      //  If a match is found, we can search it for its "Id", and set the vetTrueID accordingly
      var xmlhttp = new XMLHttpRequest();
      var url = `https://hvaf.sharepoint.com/_api/web/lists/GetByTitle('WebForm Responses')/items?$filter=Title eq '${vetTitleID}'&$top=1`;
  
      xmlhttp.onreadystatechange = function () {
        //  If successful
        if (this.readyState == 4 && this.status == 200) {
          
          //  The response is in XML, so we will need to be able to parse it. 
          let parser = new DOMParser();      
          let xmlDoc = parser.parseFromString(this.responseText, "text/xml");
          
          try {
            //  We search for the "Id"
            console.log(xmlDoc.getElementsByTagName("d:Id")[0].childNodes[0].nodeValue);
          } catch (error) {
            //  If we get an error, we know that the "ID - Real" was invalid
            document.querySelector("#vetRecordTextFeild").innerHTML = `Error: Not a valid Item ID`;
            vetTitleID = -1;            
            return false;
          }

          //  If we've gotten this far, we know that the "Id" is valid, so we can stick it in vetTrueID (but first we need to cast it as a number)
          var vetTrueIDTEMP : number = parseFloat(xmlDoc.getElementsByTagName("d:Id")[0].childNodes[0].nodeValue);
          vetTrueID = vetTrueIDTEMP;          
          
          //  Now, we'll take advantage of the API call we just made to translate to the "Id" (which was nessesary) to present the Front Desk personell with a quick
          //  info dump on the vet (so that they can double check the critical info and be able to quickly identify the vet)
          try {
            const txtBox = document.querySelector("#vetRecordTextFeild");
            txtBox.innerHTML = `
            ${xmlDoc.getElementsByTagName("d:FirstName")[0].childNodes[0].nodeValue} ${xmlDoc.getElementsByTagName("d:LastName")[0].childNodes[0].nodeValue} 
            ${xmlDoc.getElementsByTagName("d:Ethnicity")[0].childNodes[0].nodeValue} ${xmlDoc.getElementsByTagName("d:Gender")[0].childNodes[0].nodeValue} 
            Branch of Service: ${xmlDoc.getElementsByTagName("d:BranchofService")[0].childNodes[0].nodeValue} &#13;&#10;
            ${xmlDoc.getElementsByTagName("d:ServicesRequested")[0].childNodes[0].nodeValue} `;
  
          } catch (error) {
            //  If this errors out, it is likely because the vet's record is imcomplete, and that the information simply doesn't exist. Since this info
            //  is nessesary, we tell the Fron Desk to go back to the WebForm Responses List to look for an empty feild or incorrectly entered info
            const txtBox = document.querySelector("#vetRecordTextFeild");
            txtBox.innerHTML = "Something went wrong while trying to retrive this vet's info. It's most likely a problem with the way the Vet Info was entered on the GSR! You can look for and correct the error manually on the WebForm Responses List.";
            vetTitleID = -1;  
          }
          
        }
      };
      
      //  The actual API call. The above defines what do upon a response, these calls actually send the request
      xmlhttp.open("GET", url, true);
      xmlhttp.send();
    }

    //  Simply checks if the TitleID is still in the "Error State" (-1), and, if so, tells the user to Check the Record first
    function isIDValid() {
     if (vetTitleID == -1) {
      document.querySelector("#vetRecordTextFeild").innerHTML = `Error: You must first enter a valid ID and hit "Check Record"`;
      return false;
     }
     return true;
    }
    
  } 

  public _updateItemAssign(value: number): void {
    
    //  The body defines which elements of a record we want to change. 122 is the code for the "Volunteers" group. It is stored as a string and a number, so we need to replace both
    let body = '{"__metadata":{"type":"SP.Data.WebForm_x0020_ResponsesListItem"},"AssignedToStringId":"122","AssignedToId":122}'; 
    //  spHttpClient is a prebuilt framework for calling the SharePoint API. It handles the authentication headers for us
    this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('WebForm Responses')/items(${value})`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=verbose',
          'odata-version': '',
          'IF-MATCH': '*',
          'X-HTTP-Method': 'MERGE'
        },
        body: body
      })
      .then((response: SPHttpClientResponse): void => {
        //  Once the promise is fufilled and the API call has gone through (successfully)
        //  we overwrite the Modal's "display: none;" style, essentially making it "appear" (even thought it was there all along, just hidden). The Modal is on the highest z-index,
        //  so it prevents the user from clicking underneath it. The only button available is the refresh button, which sets us up for another use, no resetitng variables required!
        document.getElementById("myModal").style.display = "block";  
      }, (error: any): void => {
        console.log(error);
        document.querySelector("#vetRecordTextFeild").innerHTML = `Error: Something has gone wrong with the SharePoint API call. Try refreshing the page and trying again.`;
        // To IT: it could be that the code for the "Volunteers" group has changed from "122", or that the group no longer exists
      });    

  }
  //  End of Front Desk Logic
  
  

  //  Volunteer Logic
  private _captureClickVolunteer(): void{

     //  Sets up an asynchonous function which inserts a refresh button after 300,000ms, (5min)
     setTimeout(() => { document.getElementById("addRefresh2").innerHTML = '<div>Some information may be out of date. Click here to refresh! <button onClick="window.location.reload();">Refresh Page</button></div>'; }, 300000);

    //  Add a click event listener to the Check Record Button, calling a function which handles the input upon the click
    document.querySelector("#btnClickCheck2").addEventListener("click", handleClickCheck2);
    
    document.querySelector("#btnClickSubmit2").addEventListener("click", (e:Event) => {
      if (isIDValid2()) {
        //  Updates the Item using the TrueID
        this._updateItemRecord(vetTrueID);
      }
    });

    //  vetTitleID is the "ID - Real" of a particular entry. SharePoint uses a globally unique "Id" for API calls though, so we must differentiate (and eventually translate) between the two
    //  We set each to -1 to preform validity checks later on. If translation fails, for example, vetTrueID will remain negative, and will throw an error
    var vetTitleID = -1;
    var vetTrueID = -1;
  
    

    function handleClickCheck2() {
      //  TypeScript requires we cast certain elements as "HTML" elements, so we set up a temporary landing zone for one such value (the input for the "Check Record" button, i.e. the "ID - Real")
      var vetTitleIDTEMP : number = parseFloat((<HTMLInputElement>document.getElementById("txtItemID2")).value);
      vetTitleID = vetTitleIDTEMP;

      //  In case the input is blank (i.e. the user has not yet entered a value), we need to throw an error. If the input is blank, the above will have thrown an error, so we must now 
      //  grab the value as a string
      var vetTitleIDSTRING = (<HTMLInputElement>document.getElementById("txtItemID2")).value;
      
      //  If the string is empty, we throw an error
      if (vetTitleIDSTRING == '') {
        document.querySelector("#vetRecordTextFeild2").innerHTML = `Error: You must enter an Item ID`;
        vetTitleID = -1;
        return false;
      }

      //  If we've gotten this far, we can now begin the translation between the "ID - Real" aka vetTitleID, and the "Id" aka vetTrueID
      //  We do this by making an API call filtering the items by title, and search for entries that match our vetTitleID.
      //  If a match is found, we can search it for its "Id", and set the vetTrueID accordingly
      var xmlhttp = new XMLHttpRequest();
      var url = `https://hvaf.sharepoint.com/_api/web/lists/GetByTitle('WebForm Responses')/items?$filter=Title eq '${vetTitleID}'&$top=1`;
  
      xmlhttp.onreadystatechange = function () {
        //  If successful
        if (this.readyState == 4 && this.status == 200) {
          
          //  The response is in XML, so we will need to be able to parse it. 
          let parser = new DOMParser();
          let xmlDoc = parser.parseFromString(this.responseText, "text/xml");
          
          try {
            //  We search for the "Id"
            console.log(xmlDoc.getElementsByTagName("d:Id")[0].childNodes[0].nodeValue);         
          } catch (error) {
            //  If we get an error, we know that the "ID - Real" was invalid
            document.querySelector("#vetRecordTextFeild2").innerHTML = `Error: Not a valid Item ID`;
            vetTitleID = -1;
            return false;
          }

          //  If we've gotten this far, we know that the "Id" is valid, so we can stick it in vetTrueID (but first we need to cast it as a number)
          var vetTrueIDTEMP : number = parseFloat(xmlDoc.getElementsByTagName("d:Id")[0].childNodes[0].nodeValue);
          vetTrueID = vetTrueIDTEMP;

          //  Now, we'll take advantage of the API call we just made to translate to the "Id" (which was nessesary) to present the Volunteer with a quick
          //  info dump on the vet (so that they can double check the critical info and be able to quickly identify the vet)
          try {
            const txtBox = document.querySelector("#vetRecordTextFeild2");
            txtBox.innerHTML = `
            ${xmlDoc.getElementsByTagName("d:FirstName")[0].childNodes[0].nodeValue} ${xmlDoc.getElementsByTagName("d:LastName")[0].childNodes[0].nodeValue} 
            ${xmlDoc.getElementsByTagName("d:Ethnicity")[0].childNodes[0].nodeValue} ${xmlDoc.getElementsByTagName("d:Gender")[0].childNodes[0].nodeValue} 
            Branch of Service: ${xmlDoc.getElementsByTagName("d:BranchofService")[0].childNodes[0].nodeValue} &#13;&#10;
            Food Points Available: ${xmlDoc.getElementsByTagName("d:FoodPoints")[0].childNodes[0].nodeValue}
            Hygiene Points Available: ${xmlDoc.getElementsByTagName("d:HygienePoints")[0].childNodes[0].nodeValue} `;
            
            //  CURRENTLY NOT WORKING. NOT SURE WHY, BUT NOT CRITCAL 
            //  Sets the maxFoodPoints variable, which is used in the HTML render to limit the value the up button can go to
            maxFoodPoints = parseInt(xmlDoc.getElementsByTagName("d:FoodPoints")[0].childNodes[0].nodeValue);
          
          } catch (error) {
             //  If this errors out, it is likely because the vet's record is imcomplete, and that the information simply doesn't exist.
            const txtBox = document.querySelector("#vetRecordTextFeild");
            txtBox.innerHTML = "Something went wrong while trying to retrive this vet's info. It's most likely a problem with the way the Vet Info was entered on the GSR! Please let the front desk know.";
            vetTitleID = -1;  
          }
                   
        }
      };
   
      //  The actual API call. The above defines what do upon a response, these calls actually send the request
      xmlhttp.open("GET", url, true);
      xmlhttp.send();
    }

    //  Simply checks if the TitleID is still in the "Error State" (-1), and, if so, tells the user to Check the Record first    
    function isIDValid2() {
     if (vetTitleID == -1) {
      document.querySelector("#vetRecordTextFeild2").innerHTML = `Error: You must first enter a valid ID and hit "Check Record"`;
      console.log("Error: User attempted to assign a nonexistant item");
      return false;
     }
     return true;
    }
  
  }

  public _updateItemRecord(value: number): void{

    //  foodVal is set to the value of the Food Points button, as a number
    const foodVal = document.querySelector("#Row1Val");
    var foodPointsUsed = parseInt(foodVal.innerHTML);

    //  hygieneVal is set to the value of the Hygiene Points button, as a number
    const hygieneVal = document.querySelector("#Row2Val");
    var hygienePointsUsed = parseInt(hygieneVal.innerHTML);

    //  foodLbsRegistered is set to the value of the lbsFoodVal input, as a number
    var foodLbsRegistered : number = parseFloat((<HTMLInputElement>document.getElementById("lbsFoodVal")).value);
    
    //  clothesLbsRegistered is set to the value of the lbsClothesVal input, as a number    
    var clothesLbsRegistered : number = parseFloat((<HTMLInputElement>document.getElementById("lbsClothesVal")).value);

    //  Grabs the current time, for billing purposes
    var d = new Date();
    var timeFinished : string = d.toISOString();

    //  The body defines which elements of a record we want to change. Since our updated valus are variables, we'll need to Stringify them 
    //  instead of passing them directly as a hard coded string
    let body: string = JSON.stringify({
      '__metadata': { 'type': 'SP.Data.WebForm_x0020_ResponsesListItem' }, 'FoodPointsUsedonVisit':foodPointsUsed, 'HygienePointsUsedonVisit':hygienePointsUsed, 'Status':'Closed', 'Lbs_x002e_Food':foodLbsRegistered, 'Lbs_x002e_Clothes':clothesLbsRegistered, 'TimeVolunteerDone':timeFinished
    });
    this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('WebForm Responses')/items(${value})`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=verbose',
          'odata-version': '',
          'IF-MATCH': '*',
          'X-HTTP-Method': 'MERGE'
        },
        body: body
      })
      .then((response: SPHttpClientResponse): void => {
        //  Once the promise is fufilled and the API call has gone through (successfully)
        //  we overwrite the Modal's "display: none;" style, essentially making it "appear" (even thought it was there all along, just hidden). The Modal is on the highest z-index,
        //  so it prevents the user from clicking underneath it. The only button available is the refresh button, which sets us up for another use.
        document.getElementById("myModal2").style.display = "block";  

        //  Just in case, since any sort of accidental double submission would wipe the previous data, we clear the inputs
        document.querySelector("#Row1Val").innerHTML = '';
        document.querySelector("#Row2Val").innerHTML = '';
        (<HTMLInputElement>document.getElementById("lbsClothesVal")).value = "";
        (<HTMLInputElement>document.getElementById("lbsFoodVal")).value = "";
        (<HTMLInputElement>document.getElementById("txtItemID2")).value = "";
        
      }, (error: any): void => {
        console.log(error);
        document.querySelector("#vetRecordTextFeild").innerHTML = `Error: Something has gone wrong with the SharePoint API call. Try refreshing the page and trying again.`;
      });

  }
  //  End of Volunteer Logic

}

/////////////////////////////////////////////////////////////////////////
//  TO-DO
//    1. Bin Tracker
//    2. More Aesthietic improvments for Ipad Users
//    3. Make MaxFoodPoints acutally apply to the HTML
/////////////////////////////////////////////////////////////////////////
//    One Extra Line so that we're not at 666