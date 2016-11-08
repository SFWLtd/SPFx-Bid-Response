import * as React from 'react';
import { css } from 'office-ui-fabric-react';
import {
  Button, Label, TextField, DatePicker, Toggle
} from 'office-ui-fabric-react/lib/index';
import '../NewbidCSS.scss';
import { INewbidWebPartProps } from '../INewbidWebPartProps';
import * as $ from "jquery";

export interface INewbidProps extends INewbidWebPartProps {
}

const DayPickerStrings = {
  months: [
    'January',
    'February',
    'March',
    'April',
    'May',
    'June',
    'July',
    'August',
    'September',
    'October',
    'November',
    'December'
  ],

  shortMonths: [
    'Jan',
    'Feb',
    'Mar',
    'Apr',
    'May',
    'Jun',
    'Jul',
    'Aug',
    'Sep',
    'Oct',
    'Nov',
    'Dec'
  ],

  days: [
    'Sunday',
    'Monday',
    'Tuesday',
    'Wednesday',
    'Thursday',
    'Friday',
    'Saturday'
  ],

  shortDays: [
    'S',
    'M',
    'T',
    'W',
    'T',
    'F',
    'S'
  ],

  goToToday: 'Go to today'
};
// Creates a folder structure within a document library, inserting in the meta data specified
var createFolderStandard = function (listTitle, folderUrl, folderUrl2, bidDeadline, salesLead, teamLead, projectManager) {
  var ctx = SP.ClientContext.get_current();
  var list = ctx.get_web().get_lists().getByTitle(listTitle);

  var createFolderInternal = function (parentFolder, folderUrl) {
    // Use of promises as the async query will sometimes fail when creating additional sub folders
    var dfd = $.Deferred(function () {
      var ctx = parentFolder.get_context();
      var folderNames = folderUrl.split('/');
      var folderName = folderNames[0];
      var curFolder = parentFolder.get_folders().add(folderName);
      ctx.load(curFolder);
      console.log('1' + ' Folder URL: ' + folderUrl + ' Folder Name: ' + folderName);

      var folderItem = curFolder.get_listItemAllFields();
      folderItem.set_item("Bid_x0020_Deadline", bidDeadline);
      folderItem.set_item("Sales_x0020_Lead", salesLead);
      folderItem.set_item("Team_x0020_Lead", teamLead);
      folderItem.set_item("Project_x0020_Manager", projectManager);
      folderItem.update();
      ctx.load(curFolder);

      ctx.executeQueryAsync(
        function () {
          if (folderNames.length > 1) {
            var subFolderUrl = folderNames.slice(1, folderNames.length).join('/');
            createFolderInternal(curFolder, subFolderUrl);
            dfd.resolve();
          }
        },

        function () {
          alert('No');
          dfd.reject();
        }
      )
    });
    return dfd.promise();
  };

  var folderStatus = createFolderInternal(list.get_rootFolder(), folderUrl);
  folderStatus.done(function () {
    createFolderInternal(list.get_rootFolder(), folderUrl2);
  });
};
// Identical function to the previous function except it only creates the root folder of the bid
var createFolderNonStandard = function (listTitle, folderUrl, bidDeadline, salesLead, teamLead, projectManager) {
  // Promises no longer needed as there is only 1 folder being created
  var ctx = SP.ClientContext.get_current();
  var list = ctx.get_web().get_lists().getByTitle(listTitle);

  var createFolderInternal = function (parentFolder, folderUrl) {

    var ctx = parentFolder.get_context();
    var folderNames = folderUrl.split('/');
    var folderName = folderNames[0];
    var curFolder = parentFolder.get_folders().add(folderName);
    ctx.load(curFolder);
    console.log('1' + ' Folder URL: ' + folderUrl + ' Folder Name: ' + folderName);

    var folderItem = curFolder.get_listItemAllFields();
    folderItem.set_item("Bid_x0020_Deadline", bidDeadline);
    folderItem.set_item("Sales_x0020_Lead", salesLead);
    folderItem.set_item("Team_x0020_Lead", teamLead);
    folderItem.set_item("Project_x0020_Manager", projectManager);
    folderItem.update();
    ctx.load(curFolder);

    ctx.executeQueryAsync(
      function () {
        if (folderNames.length > 1) {
          var subFolderUrl = folderNames.slice(1, folderNames.length).join('/');
          createFolderInternal(curFolder, subFolderUrl);

        }
      },

      function () {
        alert('No');
      }
    )

  };

  createFolderInternal(list.get_rootFolder(), folderUrl);

};

// Attempt at creating a function that will query the Organisation list to retrieve a field from every item

function retrieveListItems() {

  var clientContext = SP.ClientContext.get_current();
  var oList = clientContext.get_web().get_lists().getByTitle('Organisation');

  var camlQuery = new SP.CamlQuery();
  camlQuery.set_viewXml('<Query><OrderBy><FieldRef Name="Market_x0020_Sector" /></OrderBy></Query>');
  var collListItem = oList.getItems(camlQuery);

  clientContext.load(collListItem);

  clientContext.executeQueryAsync(onQuerySucceededRL, onQueryFailedRL);

}

function onQuerySucceededRL(sender, args) {

  var clientContext = SP.ClientContext.get_current();
  var oList = clientContext.get_web().get_lists().getByTitle('Organisation');
  var camlQuery = new SP.CamlQuery();
  camlQuery.set_viewXml('<Query><OrderBy><FieldRef Name="Market_x0020_Sector" /></OrderBy></Query>');
  var collListItem = oList.getItems(camlQuery);

  var listItemInfo = '';

  var listItemEnumerator = collListItem.getEnumerator();

  while (listItemEnumerator.moveNext()) {
    var oListItem = listItemEnumerator.get_current();
    listItemInfo += '\nID: ' + oListItem.get_id() +
      '\nTitle: ' + oListItem.get_item('Title') +
      '\nBody: ' + oListItem.get_item('Body');
  }

  alert(listItemInfo.toString());
}

function onQueryFailedRL(sender, args) {

  alert('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
}

export default class Newbid extends React.Component<INewbidProps, {}> {

  public isChecked = false;
  public results;

  // Default ID's for text fields used as Fabric UI does not override the ID field displayed on the page, executed on Create Bid button

  public executeStuff() {
    var output;
    if (this.runValidation(output) == 'Clear') {
      if (this.isChecked) {
        var folderUrl1 = $('#TextField3').val() + "/1.1";
        var folderUrl2 = $('#TextField3').val() + "/1.2";
        createFolderStandard("Bids", folderUrl1, folderUrl2, $('#TextField11').val(), $('#TextField13').val(), $('#TextField15').val(), $('#TextField17').val());
      }
      else {
        var nonStandardFolderUrl = $('#TextField3').val();
        createFolderNonStandard("Bids", nonStandardFolderUrl, $('#TextField11').val(), $('#TextField13').val(), $('#TextField15').val(), $('#TextField17').val());
      }
      $(".containerStart").hide();
      $(".containerEnd").show();
      $(".createButton").show();
      $(".errorMessage").html('');
    }
    else {
      $(".errorMessage").html(this.runValidation(output));
    }
  }

  public runValidation(errorMessage) {
    if ($('#TextField3').val() == '') {
      errorMessage = "Organisation must be specified";
      return errorMessage;
    }
    else if ($('#TextField11').val() == '') {
      errorMessage = "Bid Deadline must be specified";
      return errorMessage;
    }
    else if ($('#TextField13').val() == '') {
      errorMessage = "Sales Lead must be specified";
      return errorMessage;
    }
    else if ($('#TextField15').val() == '') {
      errorMessage = "Team Lead must be specified";
      return errorMessage;
    }
    else if ($('#TextField17').val() == '') {
      errorMessage = "Project Manager must be specified";
      return errorMessage;
    }
    else {
      errorMessage = "Clear";
      return errorMessage;
    }
  }

  public handleChecked() {
    if (this.isChecked) {
      this.isChecked = false;
    }
    else {
      this.isChecked = true;
    }
  }

  public resetForm() {
    location.reload();
  }

  public render(): JSX.Element {
    return (
      <div className='container'>
        <Button className='createButton' onClick={() => this.resetForm()}>
          Create another bid?
        </Button>
        <div className='containerStart'>
          <Label className='bidLabel'>New Bids</Label>
          <div className='dataInput'>
            <div className='divDropdown'>
              <TextField className='orgSearch' label='Organisation' />
              <Button onClick={() => retrieveListItems()} >
                Search
              </Button>
              <Button
                href='https://jimboslowbro.sharepoint.com/sites/dev/SitePages/neworganisation.aspx'>
                New Organisation
              </Button>
            </div>
            <DatePicker strings={DayPickerStrings} placeholder='Bid Deadline' />
            <TextField className='textSales' label='Sales Lead' />
            <TextField className='textTeam' label='Team Lead' />
            <TextField className='textProject' label='Project Manager' />
            <Toggle label='Use standard folder structure' onChanged={() => this.handleChecked()} onText='Yes' offText='No' />
            <div className='divButton'>
              <Button
                icon='Add'
                className='orgButton'
                onClick={() => this.executeStuff()} >
                Create Bids
            </Button>
            </div>
          </div>
          <Label className='errorMessage'>
          </Label>
        </div>
        <div className='containerEnd'>
          Your bid has been submitted
        </div>
      </div>
    );
  }
}