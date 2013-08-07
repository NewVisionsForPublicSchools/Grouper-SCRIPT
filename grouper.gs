var scriptTitle = "Grouper Script V1.4 (4/18/13)";
// Written by Andrew Stillman for New Visions for Public Schools
// Published under GNU General Public License, version 3 (GPL-3.0)
// See restrictions at http://www.opensource.org/licenses/gpl-3.0.html
var GROUPERICONURL = 'https://c04a7a5e-a-3ab37ab8-s-sites.googlegroups.com/a/newvisions.org/data-dashboard/searchable-docs-collection/grouperIcon.gif?attachauth=ANoY7cr-g0qe-7Qknzq3buuhyDrSbaw27S55vEI7xNJfEuh-vNj5ruXq6Bh-jFWKMSdIWHc3NNeATWfV9eZlDtyJoVOmTloXtuwRQgWKFv8lYTZ1_GmFPFX4rFgBErqbwozJlzJauXrtMsBLJg-5bY-ukaNVJhz6Iei-xm2KZXY-nTgzUdLE1jT0nuGpDwu7ud6GPcJ9ctvKxEeTRsieiwQySq3udd3eDMvDvw_2-PKPgILejBVXcdlLTQAHqSRbpBo91gTWxBtG&attredirects=0';

// Menu unfolds differently for the owner vs. for collaborators (other editors)
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var owner = ss.getOwner().getEmail().toLowerCase();
  var user = Session.getActiveUser().getEmail().toLowerCase();
  var menuItems = [];
  var initConfig = ScriptProperties.getProperty('initConfig');
  var selectStep = ScriptProperties.getProperty('selectStep');
  menuItems[0] = {name:'What is Grouper?',functionName:'grouper_whatIs'};
  if (initConfig) {
    if (owner==user) {
      menuItems.push({name:'Review setup instructions',functionName:'grouper_howToPublishAsWebApp'});
    }
    menuItems.push(null);
    if (owner==user) {
      menuItems.push({name:'Select/delegate groups to manage',functionName:'groupPicker'}); 
      if (selectStep) {
        menuItems.push(null);
        menuItems.push({name:'Download ALL group lists',functionName:'refreshGroupSheets'});
        menuItems.push({name:'Download THIS group list',functionName:'refreshThisSheet'});
      }
    } else {
      if (selectStep) {
        menuItems.push({name:'Download this group list',functionName:'doGetRefresh'});
      }
    }
    menuItems.push(null);
    if ((owner==user)&&(selectStep)) {
      menuItems.push({name:'Upload ALL group lists',functionName:'loadUpdates'});
      menuItems.push({name:'Upload THIS group list',functionName:'updateThisSheet'});
    } else {
      if (selectStep) {
        menuItems.push({name:'Upload this group list',functionName:'doGetUpdate'});
      }
    }
   
  } else {
    menuItems.push({name:'Setup instructions',functionName:'grouper_howToPublishAsWebApp'});
  }
  ss.addMenu('Grouper', menuItems);
}


function groupPicker() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var app = UiApp.createApplication().setHeight(550);
  app.setTitle("Select the groups you want to manage.");
  app.add(app.createLabel('Optionally delegate group management other users.'));
  var panel = app.createVerticalPanel().setStyleAttribute('margin', '10px').setHeight("400px");
  try {
    var groups = GroupsManager.getAllGroups();
  } catch(err) {
    Browser.msgBox("It would appear you are trying to set up Grouper on a Gmail account or a domain account without super-admin rights.");
  }
  var groupSelections = [];
  for (var i=0; i<groups.length; i++) {
    groupSelections.push(groups[i].getName() + " - " + groups[i].getId());
  }
  var topGrid = app.createGrid(1, 2);
  topGrid.setWidget(0, 0, app.createLabel("Group Name").setWidth("180px").setStyleAttribute("marginLeft","30px").setStyleAttribute('padding', '3px')).setStyleAttribute(0, 0, 'backgroundColor', '#e5e5e5');
  topGrid.setWidget(0, 1, app.createLabel("Assign List Editing Rights (email address(es), comma separated)").setStyleAttribute('padding', '3px')).setStyleAttribute(0, 1, 'backgroundColor', '#e5e5e5');
  var addPanel = app.createHorizontalPanel().setHeight("100px");
  var addButton = app.createButton("+").setWidth("50px").setId('addButton');
  var autoComplete = SuggestBoxCreator.createSuggestBox(app,'newGroupsSelected', 350, groupSelections);  
  addPanel.add(autoComplete).add(addButton);
  app.add(topGrid);
  var scrollPanel = app.createScrollPanel().setWidth("100%").setHeight("250px").setStyleAttribute('backgroundColor', 'whiteSmoke');
  var grid = app.createGrid().setId('grid').setStyleAttribute('marginTop', '4px');
  panel.add(grid);
  var numGroups = app.createHidden('numGroups').setId('numGroups').setValue(0);
  var existingGroupIdValue = app.createHidden('existingGroupIds').setId('existingGroupIdValues');
  panel.add(numGroups);
  scrollPanel.add(panel);
  app.add(scrollPanel);
  app.add(app.createLabel("Begin typing group name(s) and use the + to add groups to the list"));
  app.add(addPanel);
  panel.add(existingGroupIdValue);
  var handler = app.createServerHandler('saveGroupSettings').addCallbackElement(panel);
  var image = app.createImage(GROUPERICONURL).setVisible(false).setStyleAttribute('position', 'absolute').setHeight("200px").setWidth("200px").setStyleAttribute('left', '100px').setStyleAttribute('top', '80px');
  var waitingHandler = app.createClientHandler().forTargets(panel).setStyleAttribute('opacity', '0.5').forTargets(image).setVisible(true);
  var button = app.createButton("Save settings", handler).addClickHandler(waitingHandler);
  app.add(button);
  app.add(image);
  if (!isWebApp()) {
    app.add(app.createLabel("Note: You have not enabled this Grouper's ability to delegate permissions via sheet editing rights. Revisit the setup instructions in the menu if you want to change this.").setStyleAttribute('padding', '10px').setStyleAttribute('backgroundColor', '#8FBC8F'));
  } else {
    app.add(app.createLabel("Note: You have published this instance of Grouper as a web app. Users you add to groups in this dialog will be granted editing rights on designated group roster sheets, allowing them manage group membership and roles from the Grouper menu in this spreadsheet.").setStyleAttribute('padding', '10px').setStyleAttribute('backgroundColor', '#FFC0CB'));
  }
  var refreshHandler = app.createServerHandler('addGroups').addCallbackElement(panel).addCallbackElement(addPanel);
  addButton.addClickHandler(refreshHandler);
  refreshGroupSettings();
  ss.show(app);
  return app;
}



function refreshGroupSettings() {
  var app = UiApp.getActiveApplication();
  var grid = app.getElementById('grid'); 
  var numGroups = app.getElementById('numGroups');
  var ssId = ScriptProperties.getProperty('ssId');
  var groupsSelected = CacheService.getPrivateCache().get('groupsSelected');
  if ((!groupsSelected)||(groupsSelected=='')) {
    groupsSelected = UserProperties.getProperty('groupsSelected||'+ssId);
  }
  var allGroupIds = [];
  var hidden = [];
  var managerBoxes = [];
  var removeHandlers = [];
  var removeButtons = [];
  var indices = [];
  var removeFlags = [];
  var i=0;
  var num = 0;
  if ((groupsSelected)&&(groupsSelected!='')) {
    groupsSelected = groupsSelected.split("||");
    grid.resize(groupsSelected.length, 6).setBorderWidth(0).setCellSpacing(0);
    for (i=0; i<groupsSelected.length; i++) {
      groupsSelected[i] = Utilities.jsonParse(groupsSelected[i]);
      hidden[i] = app.createHidden("group-"+i).setValue(groupsSelected[i].groupId);
      allGroupIds.push(groupsSelected[i].groupId);
      indices[i] = app.createHidden('index').setValue(i);
      allGroupIds.push(groupsSelected[i].groupId);
      managerBoxes[i] = app.createTextBox().setName('manager-'+i).setId('managerBox-'+i).setStyleAttribute('marginLeft', '50px').setWidth("190px").setEnabled(false);
      if (isWebApp()) {
          managerBoxes[i].setEnabled(true);
      }
      if ((groupsSelected[i].manager)&&(groupsSelected[i].manager!='')) {
        managerBoxes[i].setValue(groupsSelected[i].manager);
      }
      removeHandlers[i] = app.createServerHandler('removeGroup').addCallbackElement(indices[i]).addCallbackElement(numGroups);
      removeButtons[i] = app.createButton("X", removeHandlers[i]);
      removeFlags[i] = app.createHidden("removeFlag-"+i).setId("removeFlag-"+i).setValue(0);
      var groupName = GroupsManager.getGroup(groupsSelected[i].groupId).getName();
      grid.setWidget(i, 0, app.createLabel(groupName + " - " + groupsSelected[i].groupId));
      grid.setWidget(i, 1, hidden[i]);
      grid.setWidget(i, 2, managerBoxes[i]);
      grid.setWidget(i, 3, indices[i]);
      grid.setWidget(i, 4, removeButtons[i]);
      grid.setWidget(i, 5, removeFlags[i]);
      num++;
      numGroups.setValue(num);
    }
    app.getElementById('existingGroupIdValues').setValue(allGroupIds.join(","));
    app.getElementById('newGroupsSelected').setValue('');
  }
  return app;
}


function isWebApp() {
  var enabled = ScriptApp.getService().isEnabled();
  var url = ScriptApp.getService().getUrl();
  if ((enabled)&&(url)) {
    return true;
  } else {
    return false;
  }
}


function addGroups(e) {
  var app = UiApp.getActiveApplication();
  var grid = app.getElementById('grid');
  var ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  var newGroupIds = e.parameter.newGroupsSelected;
  newGroupIds = newGroupIds.split(",");
  var existingGroupIds = e.parameter.existingGroupIds;
  if (existingGroupIds!='') {
    existingGroupIds = existingGroupIds.split(",");
  } else {
    existingGroupIds = [];
  }
  var numGroups = app.getElementById('numGroups');
  var existingGroupIdValues = app.getElementById('existingGroupIdValues');
  var num = parseInt(e.parameter.numGroups);
  var groupsSelected = [];
  var allGroupIds = [];
  if (existingGroupIds.length>0) {
    allGroupIds = existingGroupIds.slice(0);
  } else {
    allGroupIds = [];
  }
  var i=num;
  if (newGroupIds) {
    for (var j=0; j<newGroupIds.length-1; j++) {
      if (allGroupIds.indexOf(newGroupIds[j].split(" - ")[1])==-1) {
        allGroupIds.push(newGroupIds[j].split(' - ')[1]);
        grid.resize(i+j+1, 6);
        var hidden = app.createHidden("group-"+(i+j)).setValue(newGroupIds[j].split(" - ")[1]);
        var index = app.createHidden('index').setValue(i+j);
        var managerBox = app.createTextBox().setName('manager-'+(i+j)).setId('managerBox-'+(i+j)).setStyleAttribute('marginLeft', '50px').setWidth("190px").setEnabled(false);
        if (isWebApp()) {
          managerBox.setEnabled(true);
        }
        var removeHandler = app.createServerHandler('removeGroup').addCallbackElement(index).addCallbackElement(numGroups);
        var removeButton = app.createButton("X", removeHandler);
        var removeFlags = app.createHidden("removeFlag-"+(i+j)).setId("removeFlag-"+(i+j)).setValue(0);
        grid.setWidget(i+j, 0, app.createLabel(newGroupIds[j]));
        grid.setWidget(i+j, 1, hidden);
        grid.setWidget(i+j, 2, managerBox);
        grid.setWidget(i+j, 3, index);
        grid.setWidget(i+j, 4, removeButton);
        grid.setWidget(i+j, 5, removeFlags);
        num++;
        numGroups.setValue(num);
        existingGroupIdValues.setValue(allGroupIds.join(","));
        checkSelectedGroupNum(e);
      }
    }
  }
  app.getElementById('newGroupsSelected').setValue('');
  return app;
}



function removeGroup(e) {
  var app = UiApp.getActiveApplication();
  var grid = app.getElementById('grid');
  var index = e.parameter.index;
  var num = e.parameter.numGroups;
  var removeFlag = app.getElementById('removeFlag-'+index)
  for (var i=0; i<num; i++) {
    if (i==index) {
      grid.setStyleAttribute(i, 0, 'opacity','0.3').setStyleAttribute(i, 0, 'backgroundColor','pink');
      grid.setStyleAttribute(i, 1, 'opacity','0.5').setStyleAttribute(i, 1, 'backgroundColor','pink');
      grid.setStyleAttribute(i, 2, 'opacity','0.5').setStyleAttribute(i, 2, 'backgroundColor','pink');
      grid.setStyleAttribute(i, 3, 'opacity','0.5').setStyleAttribute(i, 3, 'backgroundColor','pink');
      removeFlag.setValue(1);
    }
  }
  return app;
}


function checkSelectedGroupNum(e) {
  var app = UiApp.getActiveApplication();
  var addButton = app.getElementById('addButton');
  var num = e.parameter.numGroups;
  var j = 0;
  for (var i=0; i<num; i++) {
    if (e.parameter['group-'+i]) {
      j++;
    }
  }
  if (j>30) {
    addButton.setEnabled(false);
    var ssId = ScriptProperties.getProperty('ssId');
    var ss = SpreadsheetApp.openById(ssId);
    ss.toast("For the sake of performance and reliability, it is not recommended you exceed 30 groups in a single instance of Grouper.");
  }
  return app;
}


function saveGroupSettings(e) {
  var app = UiApp.getActiveApplication();
  var ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  var numGroups = e.parameter.numGroups;
  var groupsSelected = [];
  var j=0;
  for (var i=0; i<numGroups; i++) {
    if ((e.parameter['group-'+i])&&(e.parameter['removeFlag-'+i]!=1)) {
      groupsSelected[j] = new Object();
      groupsSelected[j].groupId = e.parameter['group-'+i];
      groupsSelected[j].manager = e.parameter['manager-'+i];
      Logger.log(groupsSelected[j].groupId + " - " + groupsSelected[j].manager);
      groupsSelected[j] = Utilities.jsonStringify(groupsSelected[j]);
      j++;
    }
  }
  var newGroupsSelectedBox = app.getElementById('newGroupsSelected');
  newGroupsSelectedBox.setValue('');
  groupsSelected = groupsSelected.join("||");
  UserProperties.setProperty('groupsSelected||'+ssId, groupsSelected);
  ScriptProperties.setProperty('selectStep','true');
  app.close();
  addDeleteSheets();
  onOpen();
  return app;
}



// Ui providing step by step instructions for how to publish the script as a webApp
function grouper_howToPublishAsWebApp() {
  setGrouperUid();
  setGrouperSid();
  var app = UiApp.createApplication().setTitle('Step 1. Set up Grouper').setHeight(540).setWidth(700);
  var thisSs = SpreadsheetApp.getActiveSpreadsheet();
  ScriptProperties.setProperty('ssId', thisSs.getId());
  var panel = app.createVerticalPanel();
  var handler = app.createServerHandler('Grouper_confirmSettings').addCallbackElement(panel);
  var button = app.createButton("Thanks, I've got it!").addClickHandler(handler);
  var scrollpanel = app.createScrollPanel().setHeight("460px");
  var html = app.createHTML('Grouper is designed to be run from an account that has "super admin" rights on the domain, with the User Provisioning feature turned on in the control panel. Please visit your domain control panel and turn on User Provisioning if you wish to use this script.').setStyleAttribute('marginTop', '5px');
  panel.add(html);
  panel.add(app.createImage('https://c04a7a5e-a-3ab37ab8-s-sites.googlegroups.com/a/newvisions.org/data-dashboard/searchable-docs-collection/provisioning%20api.png?attachauth=ANoY7cpXIpgEF_-Ogxup7WW1NSa9YFJmnJXSvgyxz7NeoP-lJjQniaoZwaD_DDfo90-QtcE9vFqZ6JoFX742hFerXCBB3lfQn81l_F_6Obpavna6VxdbD_bAeJdBWr910yphend0bnUGi4N3sDlMZNZNc0KBOINdRCGe7YyxpMsLuohX9_vt8f9Ya9nfM2eKdxG7z_e2SirJg-8fBkjjktfl_HiHw6xBWE2yyDsCe1PihUk55FdaeIBXVXzee8bGgUDMvF189VuH&attredirects=0').setWidth("650px"));
  panel.add(app.createLabel("Optional: Delegate group management via Spreadsheet editing rights").setStyleAttribute("width","100%").setStyleAttribute("backgroundColor", "grey").setStyleAttribute("color", "white").setStyleAttribute("padding", "5px 5px 5px 5px"));
  var html2 = app.createHTML('If you wish to use Grouper to delegate group management rights to users who are not domain "super admins," it will require publishing this script (from a super admin account) as a web app limited to users in your domain.  Operating "Grouper" in this manner gives designated users editing rights on specific sheets in the spreadsheet, and exposes the code of this script to those same users as editors, so <strong>you should not enable this feature unless you can trust the specific users you are delegating control to.</strong> The instructions below explain how to publish this script as a web app.').setStyleAttribute('marginTop', '5px');
  panel.add(html2);
  var grid = app.createGrid(10, 2).setBorderWidth(0).setCellSpacing(0);
  var text1 = app.createLabel("Instructions:").setStyleAttribute("width", "100%").setStyleAttribute("backgroundColor", "grey").setStyleAttribute("color", "white").setStyleAttribute("padding", "5px 5px 5px 5px");
  app.add(text1);
  grid.setWidget(0, 0, app.createHTML('1. Go to \'Tools->Script editor\' from the Spreadsheet that contains your form.</li>'));
  grid.setWidget(0, 1, app.createImage('https://c04a7a5e-a-3ab37ab8-s-sites.googlegroups.com/a/newvisions.org/data-dashboard/searchable-docs-collection/grouper1.png?attachauth=ANoY7cpydoPLs25bw8pZBtgD1rH4o7xZexwhDS1EMF695O7RdWG3ULvjWTXIrdzAT5xHoa0odiuSOWd2WASUFojg-jBxbCx65gp_HfwV3YDTFC36v2beyjGpmXa-Q9KUbz2fMcthRlYIcKnuHWxABGZ5oN9n9CUIXE-f7ajQ_EDo1gAn-bbQFsHuLE2SkBwtxY_xlnOrxeyfxshI0qlKaTNdlJgMbeSfg8tyFCNBO01bVBUsyhRt-dcj2UXByr2VDpaz0-iki13z&attredirects=0').setWidth("410px"));
  grid.setStyleAttribute(1, 0, "backgroundColor", "grey").setStyleAttribute(1, 1, "backgroundColor", "grey");
  grid.setWidget(2, 0, app.createHTML('2. Under the \'File\' menu in the Script Editor, select \'Manage versions\' and save a new version of the script. Because it\'s optional, you can leave the \'Describe what changed\' field blank.</li>'));
  grid.setWidget(2, 1, app.createImage('https://c04a7a5e-a-3ab37ab8-s-sites.googlegroups.com/a/newvisions.org/data-dashboard/searchable-docs-collection/grouper2.png?attachauth=ANoY7crP7ZjKHKnGHHdGU_rzZ6o9OMAE2LRSGJKwNs8lMTQNsyLhg0GU3bExpfTUi5xwb8hJe6WkWnuIi2m90llj0qJ7BFobpLLT8zg3rhd0Utd2LsdwapCuPRCioOGAwKVD5uC7fmEe-BMYqaIVnfFjHe9AUzt-Rvql5D1I8xzgAeYnh7JdpgoJckiRYMt6c8wzwdGBx0xaVHHnG65w7aWjvd_tiP0Oe7i6-L1YU4vU9aJF6l105nGt1vNcrWZ5wf9vqS24AuC4&attredirects=0').setWidth("410px"));
  grid.setWidget(3, 1, app.createImage('https://c04a7a5e-a-3ab37ab8-s-sites.googlegroups.com/a/newvisions.org/data-dashboard/images-for-formmutant/manage%20versions.png?attachauth=ANoY7cpVfU-WEwiPuOMKMIAzbK6EdA8xkmv_M2R8GKdlcGLC7mo00ZJykbBFrtJEZQHDpKVdvizQQnuyfGVc65iigmGuGr_ZwC2Z4rnh1V67_ogOJKXH2TWmDAafxa-q_5fngrasDYYN2w2-hR_eR95GoY6e5Rza-mtWb1iAp97Cm8n9kVHRk67dURdrdD5AIaS8ZOkse1MmfaN-ZJpMv7bLYBKpisq8GldTTjo7W55OUIJhFuDcxLEc__vguXArjfb9Pd_e2bZD&attredirects=0').setWidth("410px"));
  grid.setStyleAttribute(4, 0, "backgroundColor", "grey").setStyleAttribute(4, 1, "backgroundColor", "grey");
  grid.setWidget(5, 0, app.createHTML('3. Under the \'Publish\' menu in the Script Editor, select \'Deploy as web app\'. Choose the version you want to publish (usually #1). Under \'Execute web app as\', choose \'Me (' + Session.getActiveUser().getEmail() + ')\'. Under "Who has access to the app:"  select \'anyone within ' + Session.getActiveUser().getEmail().split("@")[1] + '\''));
  grid.setWidget(5, 1, app.createImage('https://c04a7a5e-a-3ab37ab8-s-sites.googlegroups.com/a/newvisions.org/data-dashboard/searchable-docs-collection/grouper4.png?attachauth=ANoY7coRmq6s-UFXe0Odl6jgi0O8X9_FDDUz0vnPrmSXIfxcHLx83Wjt1vdY6kDJ-jgOGk56quhp3G2bmIJasHFTedfGpTNTSleSdOOqGVpYCDjPDuB8dseLaUHTDzN6XIKuc_atpszB8jdHiGhoKno_p4EzzIRuIMbvL-9Iu_l503s3msSt5gXqYmN4nzdNAH4vbjUOvNpVIMNI-B_1ryHcvGIt9JrNivUyf9nxqugDkUPr7ljn1oYjoqGaxVCDdWUHigoYvXCh&attredirects=0').setWidth("410px"));
  grid.setWidget(6, 1, app.createImage('https://c04a7a5e-a-3ab37ab8-s-sites.googlegroups.com/a/newvisions.org/data-dashboard/searchable-docs-collection/grouper5.png?attachauth=ANoY7crBtfjG9KCTn9OPlOuch4pmE7YaPav0CEce0NW1uL24-iK8k9JjLVZNJrbwjJfkHxgBBgIDtO-fp3MxiBYZIkznFD0QQGIBDQsw4mPWJyr7-U-tYtWIx66BSf2S85ufSPYhZQJUxwJiJXELgZLO-afCOXWn0VFwIuwyLbQyokofXTX96gXh_5V79WHR3H6BXPabL_8KzFGiYsv0IhBLsvXtDoGRwKrrOWUnIxGlws8_X7OnKPbpmLzXiE1U-LxQCNwnhgxD&attredirects=0').setWidth("410px"));
  grid.setStyleAttribute(7, 0, "backgroundColor", "grey").setStyleAttribute(7, 1, "backgroundColor", "grey");
  grid.setWidget(8, 0, app.createHTML('4. Once you have published the script as a web app, select the "Refresh/download all groups" menu item to create and auto-populate all sheets with your domain groups. Going forward, those who you grant editing rights on the spreadsheet will be able to maintain (add, remove members, edit roles, etc.) these groups in the sheets they have editing rights to.  It is recommended you use the sheet protection feature to restrict editing to specific editors as a method for delegating maintainer authority on specific groups.'));
  grid.setWidget(8, 1, app.createImage('https://c04a7a5e-a-3ab37ab8-s-sites.googlegroups.com/a/newvisions.org/data-dashboard/searchable-docs-collection/grouper8.png?attachauth=ANoY7coLjZAtMC3AiICKNVlCbVp1_o9rgcQ9cD4l1RD1ig8rdU0QvDoknLygH03U95qi5db5N_Hb0_fuXdAIYPZ2KbiaA0lsucBJRUrkKJHmmkUH-ZN003_HoG7Fi1GfZ0zRM-Hm-ikLm5pccjvpJsmGMcd2abnhjJCMK9OMLsgHiCc6szQ2CvfQlCmZybIoyk8FwXZOo6ISsgYZUE34Mc9UTjins263olYtRkzBC1blmyK0b6Fi8zyKRNL0lPBNStdzMrPl9ak0&attredirects=0').setWidth("410px"));
  grid.setWidget(9, 1, app.createImage('https://c04a7a5e-a-3ab37ab8-s-sites.googlegroups.com/a/newvisions.org/data-dashboard/searchable-docs-collection/grouperUi.png?attachauth=ANoY7crPO0tKhYXoTfbDsuKAi5W2YY4zEnf6ZlolFW-tU1rY-IfFbJ8DuorvQb1sFfecnvzHKdU2HvgN4g40H_wsj6If15WZX1FtUBaop-vTXRwrrFzvTHFFDmFSPDNF1YulO4p9uIhLi8BxDzZMcYLo4XsokuTnvQQuiBuxbAxipvfJK05AsdZZnaM6_7c2PZJpITbBEWC7qr48D0kRkQ8tbaxKsr9njLBYZT-Ou_QSO3O5JpPH_2-yd6xbSV8Zcy4dJ747_cqW&attredirects=0').setWidth("410px"));
  panel.add(grid);
  scrollpanel.add(panel);
  app.add(scrollpanel);
  app.add(button);
  thisSs.show(app);
  return app; 
}

function Grouper_confirmSettings(e) {
  var app = UiApp.getActiveApplication();
  ScriptProperties.setProperty('initConfig', 'true');
  onOpen();
  app.close();
  return app;
}


function Grouper_verifyPublished() {
  var scriptUrl = ScriptApp.getService().getUrl();
  if (scriptUrl) {
    return true;
  } else {
    return false;
  }
}


function refreshThisSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = ss.getActiveSheet().getName();
  refreshGroupSheets(sheetName);
}

function updateThisSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = ss.getActiveSheet().getName();
  loadUpdates(sheetName);
}

function doGetUpdate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheetName = ss.getActiveSheet().getName();
  var webAppUrl = ScriptApp.getService().getUrl();
  if (!webAppUrl) {
    Browser.msgBox("The web app for this script has not been published");
  }
  var fullUrl = webAppUrl + "?activeSheetName=" + activeSheetName + "&mode=update&ssId=" + ss.getId();
  var app = UiApp.createApplication().setHeight(100);;
  app.setTitle("Update group membership for group: " + activeSheetName);
  var closeHandler = app.createServerHandler('closeApp');
  app.add(app.createAnchor("Click to update this group", fullUrl).addClickHandler(closeHandler).setStyleAttribute('margin', '15px'));
  app.add(app.createLabel("Why this web link?  The Grouper script is designed to run as a web-app for all users who are not the spreadsheet owner.").setStyleAttribute('margin', '15px'));
  ss.show(app);
  return app;
}

function closeApp() {
  var app = UiApp.getActiveApplication();
  app.close();
  return app;
}

function doGetRefresh() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheetName = ss.getActiveSheet().getName();
  var userEmail = Session.getActiveUser().getEmail().toLowerCase(); 
  var webAppUrl = ScriptApp.getService().getUrl().toLowerCase();
  if (!webAppUrl) {
    Browser.msgBox("The web app for this script has not been published");
  }
  var fullUrl = webAppUrl + "?mode=refresh&ssId=" + ss.getId() + "&activeSheetName=" + activeSheetName;
  var app = UiApp.createApplication().setHeight(100);
  app.setTitle("Refresh/download group roster for group: " + activeSheetName);
  var closeHandler = app.createServerHandler('closeApp');
  app.add(app.createAnchor("Click to complete this action", fullUrl).addClickHandler(closeHandler).setStyleAttribute('margin', '15px'));
  app.add(app.createLabel("Why this web link?  The Grouper script is designed to run as a web-app for all users who are not the spreadsheet owner.").setStyleAttribute('margin', '15px'));
  ss.show(app);
  return app;
}

function fetchAllGroupIds () {
  var groups = GroupsManager.getAllGroups();
  var groupIds = [];
  for (var i=0; i<groups.length; i++) {
    groupIds.push(groups[i].getId());
  }
  return groupIds;
}



function addDeleteSheets() {
  var ssId = ScriptProperties.getProperty('ssId');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var message = '';
  var groupIds = [];
  var managers = [];
  var groupsSelected = UserProperties.getProperty('groupsSelected||'+ssId);
  groupsSelected = groupsSelected.split("||");
  for (var i=0; i<groupsSelected.length; i++) {
    var thisGroup = Utilities.jsonParse(groupsSelected[i]);
    groupIds.push(thisGroup.groupId); 
    managers.push(thisGroup.manager);
  }
  var sheets = ss.getSheets();
  var sheetNames = [];
  for (var i=0; i<sheets.length; i++) {
    sheetNames.push(sheets[i].getName());
  }
  for (var j=0; j<groupIds.length; j++) {
    if (sheetNames.indexOf(groupIds[j]) == -1) {
      var sheet = ss.insertSheet(groupIds[j]);
      message += "Added sheet: " + groupIds[j] + ", ";
      var permissions = sheet.getSheetProtection();
      if (managers[j]) {
        var theseManagers = managers[j].replace(/\s+/g, '').split(",");
        for (var k=0; k<theseManagers.length; k++) {
          ss.addEditor(theseManagers[k]);
          try {
            permissions.addUser(theseManagers[k]);
          } catch (err) {
          }
        }
        message += "with editor(s): " + theseManagers.join(", ");
        sheet.setColumnWidth(5, 500);
        sheet.setFrozenRows(1);
      }
      permissions.setProtected(true);
      sheet.setSheetProtection(permissions);
      refreshGroupSheets(sheet.getName());
    } else {
      var sheet = ss.getSheetByName(groupIds[j]);
      var permissions = sheet.getSheetProtection();
      var theseManagers = '';
      if ((managers[j])&&(managers[j]!='')) {
        theseManagers = managers[j].replace(/\s+/g, '').split(",");
        var sheetEditors = permissions.getUsers();
        var editors = ss.getEditors();
        var ssEditors = [];
        for (var m=0; m<editors.length; m++) {
          ssEditors.push(editors[m].getEmail());
        }
        for (var k=0; k<theseManagers.length; k++) {
          if (ssEditors.indexOf(theseManagers[k])==-1) {
            ss.addEditor(theseManagers[k]);
            message += "Added new spreadsheet editor " + theseManagers[k] + ". ";
          }
          if (sheetEditors.indexOf(theseManagers[k])==-1) {
            permissions.addUser(theseManagers[k]);
            message += "Added editor " + theseManagers[k] + " to " + groupIds[j] + ", ";
          }
        }
        for (var k=0; k<sheetEditors.length; k++) {
          if ((theseManagers.indexOf(sheetEditors[k])==-1)&&(editors[k]!=ss.getOwner().getEmail())) {
            permissions.removeUser(editors[k]);
            message += "Removed editor " + editors[k] + " from " + groupIds[j] + ", ";
          }
        }
      }
    }
    permissions.setProtected(true);
    sheet.setSheetProtection(permissions)
    sheet.getRange(1, 1, 1, 4).setValues([['Email','First Name','Last Name','Role']]);
    sheet.getRange(1, 2).setNote("Optional");
    sheet.getRange(1, 3).setNote("Optional");
    sheet.getRange(1, 4).setNote("Role must be either \"member\" or \"owner\"");
    sheet.setFrozenRows(1);
  }
  for (var i=0; i<sheets.length; i++) {
    if (groupIds.indexOf(sheets[i].getName())==-1) {
      ss.setActiveSheet(sheets[i]);
      message += "Removed sheet: " + sheets[i].getName() + ", ";
      ss.deleteActiveSheet();
    } 
  }
  if (message == '') {
    message = "No changes made.";
  }
  ss.toast(message);
  return;
}


function refreshGroupSheets(sheetName, mode) {
  var ssId = ScriptProperties.getProperty('ssId');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    ss = SpreadsheetApp.openById(ssId);
  }
  var user = Session.getActiveUser().getEmail();
  var owner = ss.getOwner().getEmail();
  var groupIds = [];
  var managers = [];
  var groupsDownloaded = '';
  var groupsSelected = UserProperties.getProperty('groupsSelected||'+ssId);
  groupsSelected = groupsSelected.split("||");
  for (var i=0; i<groupsSelected.length; i++) {
    var thisGroup = Utilities.jsonParse(groupsSelected[i]);
    groupIds.push(thisGroup.groupId); 
    managers.push(thisGroup.manager.split(","));
  }
  var index = groupIds.indexOf(sheetName);
  if ((sheetName)&&(index!=-1)&&((managers[index].indexOf(user)!=-1)||(owner==user))) {
    groupIds = [sheetName];
    var sheet = ss.getSheetByName(sheetName);
    var sheets = [sheet];
  }
  if ((sheetName)&&(managers[index].indexOf(user)==-1)&&(user!=owner)) {
    var message = "You are not authorized to update this group.";
    return message;
  }
  
  if (!sheetName) {
    var sheets = ss.getSheets();
  } else {      
  }
  var sheetNames = [];
  for (var i=0; i<sheets.length; i++) {
    if ((groupIds.indexOf(sheets[i].getName())==-1)&&(sheets[i].getName().indexOf("@")!=-1)) {
      ss.setActiveSheet(sheets[i]);
      ss.deleteActiveSheet();
    } else {
      sheetNames.push(sheets[i].getName());
    }
  }
  
  for (var j=0; j<groupIds.length; j++) {
    if (sheetNames.indexOf(groupIds[j]) == -1) {
      var sheet = ss.insertSheet(groupIds[j]);
      sheet.setColumnWidth(5, 500);
      sheet.setFrozenRows(1);
    } else {
      var sheet = ss.getSheetByName(groupIds[j]);
    }
    sheet.clearContents();
    sheet.getRange(1, 1, 1, 4).setValues([['Email','First Name','Last Name','Role']]);
    SpreadsheetApp.flush();
    var allMembers = GroupsManager.getGroup(groupIds[j]).getAllMembers();
    var owners = GroupsManager.getGroup(groupIds[j]).getAllOwners();
    var members = [];
    for (var l=0; l<allMembers.length; l++) {
      if (owners.indexOf(allMembers[l])==-1) {
        members.push(allMembers[l])
      }
    }
    var memberOwnerArray = [];
    var m = 0
    for (var k=0; k<owners.length; k++) {
      var row = memberOwnerArray[m];
      if (!row) {
        memberOwnerArray[m] = new Object();
      }
      var firstName = '';
      var lastName = '';
      try {
        firstName = UserManager.getUser(owners[k].split("@")[0]).getGivenName();
        lastName = UserManager.getUser(owners[k].split("@")[0]).getFamilyName();
      } catch(err) {
      }
      memberOwnerArray[m]['email'] = owners[k];
      memberOwnerArray[m]['firstName'] = firstName;
      memberOwnerArray[m]['lastName'] = lastName;  
      memberOwnerArray[m]['role'] = "owner";    
      m++;
    }
    
    
    for (var k=0; k<members.length; k++) {
      var row = memberOwnerArray[m];
      if (!row) {
        memberOwnerArray[m] = new Object();
      }
      var firstName = '';
      var lastName = '';
      try {
        firstName = UserManager.getUser(members[k].split("@")[0]).getGivenName();
        lastName = UserManager.getUser(members[k].split("@")[0]).getFamilyName();
      } catch(err) {
      }
      memberOwnerArray[m]['email'] = members[k];
      memberOwnerArray[m]['firstName'] = firstName;
      memberOwnerArray[m]['lastName'] = lastName;  
      memberOwnerArray[m]['role'] = "member";    
      m++;
    }
    if (memberOwnerArray.length>0) {
      setRowsData(sheet, memberOwnerArray);
    }
    sheet.sort(3, true);    
    sheet.sort(4, false); 
     if (j>0) {
      groupsDownloaded += ", ";
    }
    groupsDownloaded += groupIds[j];
    Grouper_downloadGroup();
  }
  SpreadsheetApp.flush();
  var message = "Successfully downloaded " + groupsDownloaded + " group list(s) to Spreadsheet.";
  ss.toast(message, "Roster download status");
  return message;
}


function loadUpdates(sheetName, mode) {
  var ssId = ScriptProperties.getProperty('ssId');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var message = '';
  if (!ss) {
    ss = SpreadsheetApp.openById(ssId);
  }
  if (!sheetName) {
    var sheets = ss.getSheets();
  } else {
    var sheet = ss.getSheetByName(sheetName); 
    var sheets = [sheet];
  }
  var removed = '';
  var ownerAdded = '';
  var memberAdded = '';
  for (var i=0; i<sheets.length; i++) {
    var groupId = sheets[i].getName();
    if (groupId.indexOf("@")==-1) {
      continue;
    }
    try {
      var group = GroupsManager.getGroup(groupId);
    } catch(err) {
      message = groupId + " could not be found. ";
      continue;
    }
    var allMembers = group.getAllMembers().sort();
    var owners = group.getAllOwners().sort();
    var members = [];
    for (var l=0; l<allMembers.length; l++) {
      if (owners.indexOf(allMembers[l])==-1) {
        members.push(allMembers[l])
      }
    }
    if (sheets[i].getLastRow()>1) {
      var data = getRowsData(sheets[i], sheets[i].getRange(2, 1, sheets[i].getLastRow()-1, 4))
      var listedOwners = [];
      var listedMembers = [];
      for (var j=0; j<data.length; j++) {
        if (data[j]['role']=='owner') {
          listedOwners.push(data[j]['email']);
          if (members.indexOf(data[j]['email'])!=-1) {
            group.removeMember(data[j]['email']);    
          }
          if (owners.indexOf(data[j]['email'])==-1) {
            group.addOwner(data[j]['email']);
            sheets[i].getRange(j+2, 5).setValue("Set as owner on " + new Date());
            ownerAdded += group.getName() + ": " + data[j]['email'];
            owners.push(data[j]['email']);
          }
        } 
        
        if ((data[j]['role']=='member')||(!data[j]['role'])) {
          listedMembers.push(data[j]['email']);
          if (owners.indexOf(data[j]['email'])!=-1) {
            group.removeOwner(data[j]['email']);
          }
          if (members.indexOf(data[j]['email'])==-1) {
            group.addMember(data[j]['email']);
            if (!data[j]['role']) {
              sheets[i].getRange(j+2, 4).setValue("member");
            }
            sheets[i].getRange(j+2, 5).setValue("Set as member on " + new Date());
            memberAdded += group.getName() + ": " + data[j]['email'];
            members.push(data[j]['email']);
          }
        }
      }
      for (var k=0; k<members.length; k++) {
        if (listedMembers.indexOf(members[k])==-1) {
          group.removeMember(members[k]);
          removed += group.getName() + ":" +  members[k] + " as member,";
        }
      }
      for (var k=0; k<owners.length; k++) {
        if (listedOwners.indexOf(owners[k])==-1) {
          group.removeOwner(owners[k]);
          removed += group.getName() + ":" + owners[k] + " as owner, ";
        }
      }
    }   
    Grouper_updateGroup();
  }
  SpreadsheetApp.flush();
  if ((removed!='')||(memberAdded!='')||(ownerAdded!='')) {
     message = "The groupAdmin script ";
    if (removed!='') {
      message += "removed the following: " + removed + ", ";
    }
    if (memberAdded!='') {
      message += "added the following as members: " + memberAdded + ", ";
    }
    if (ownerAdded!='') {
      message += "added the following as owners: " + ownerAdded + ", "; 
    }
  } else {
    message = "No changes were detected or made to domain groups.";
  }
  if (mode) {
    return message; 
    ss.toast(message, "Update status");
  } else {
    ss.toast(message, "Update status");
  }
}
