var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getActiveSheet();

var urlRoot = "http://zxing.org/w/chart?cht=qr&chs=350x350&chld=H&choe=UTF-8&chl=";

function include(filename) {
  var return1= HtmlService.createTemplateFromFile(filename).getRawContent();  
  return return1;
}

/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */
function onOpen(e) {
   var items = [
      {name: LanguageApp.translate('Create QR codes for contacts','',Session.getActiveUserLocale()),functionName: 'menuItemPrepararHoja'},
   ];
   ss.addMenu('Contacts2QrCode', items);
}

/**
 * DON'T AUTOLAUNCH
 *
 * Runs when the add-on is installed.
 *
 * @param {object} e The event parameter for a simple onInstall trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode. (In practice, onInstall triggers always
 *     run in AuthMode.FULL, but onOpen triggers may be AuthMode.LIMITED or
 *     AuthMode.NONE.)
 * function onInstall(e) {
 *  onOpen(e);
 * }
 */

function menuItemPrepararHoja(){        
     nextStage(1)
}

function nextStage( currentStage, contactsGroup , folder ){     
  var template = HtmlService.createTemplateFromFile('PreparePage')
     template.currentStage = currentStage;
     template.contactsGroup = contactsGroup ? contactsGroup : '';
     template.folderURL= folder ? folder.getUrl() : '';
     template.folderId = folder ? folder.getId() : '';
     template.folderName = folder ? folder.getName() : '';
     template.totalRows = sheet.getLastRow()-1;
  var html = template.evaluate();
      html.setTitle("Contacts2QRCode")
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setWidth(300);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showSidebar(html);           
}
     
function listGroups(){
  var grps = ContactsApp.getContactGroups();
  var ret = new Array();
  for(var i in grps){
    ret[i] = {
      name : grps[i].getName(),
      id : grps[i].getId()
    }
  }
  return ret;
}

function populateWithGroup(groupName){
  var grps = ContactsApp.getContactGroup(groupName);
  if( grps === null){
    return;
  }
  sheet.clearContents();
  populateHeader();
  var contacts = grps.getContacts();
  for(var i in contacts){
      var contact = contacts[i];
      populateRow(contact);    
  }
  nextStage(2, groupName);
}

function populateHeader(){
   sheet.appendRow(["Id", "Name", "JobTitle", "Phone", "eMail", "Address", "Organization", "URL", "QR"]);
}

function populateRow( contact ){
  var name = stripAccent(contact.getFullName());
  var jobTitle = "";
  var email="";
  var address="";
  var organization="";
  var phone="";
  var url=""
  for(var i in contact.getCompanies() ){
    jobTitle = stripAccent(contact.getCompanies()[i].getJobTitle());
    organization = stripAccent(contact.getCompanies()[i].getCompanyName());    
  }
  for(var i in contact.getAddresses()){
    //if( contact.getAddresses()[i].isPrimary() ){
      address = stripAccent(contact.getAddresses()[i].getAddress()).replace(/,/g,' ');
    //}
  }
  for(var i in contact.getEmails()){
    //if( contact.getEmails()[i].isPrimary() ){
      email = stripAccent(contact.getEmails()[i].getAddress());
    //}
  }
  Logger.log(contact.getPhones())
  for(var i in contact.getPhones()){
    //if( contact.getPhones()[i].isPrimary() ){
      phone = "'"+stripAccent(contact.getPhones()[i].getPhoneNumber()).replace(/\+/g,'00');
    //}
  }  
  for(var i in contact.getUrls()){
    //if( contact.getPhones()[i].isPrimary() ){
      url = "'"+stripAccent(contact.getUrls()[i].getAddress());
    //}
  }  
  var id = sheet.getLastRow();
  sheet.appendRow([id,name,jobTitle,phone,email,address,organization,url,""]);
}

function stripAccent( str ){
  var ret = ''
  str = ""+str
  Logger.log(str)  
  for(var i=0; i<str.length;i++ ){
    Logger.log(i+" === > "+str[i])  
    switch( str[i] ){
      case 'á' : ret+="%C3%A1"
      break;
      case 'é' : ret+="%C3%A9"
      break;
      case 'í' : ret+="%C3%AD"
      break;
      case 'ó' : ret+="%C3%B3"
      break;
      case 'ú' : ret+="%C3%BA"
      break;
      case 'ñ' : ret+="%C3%B1"
      break;
      case 'Á' : ret+= "%C3%81"
      break;
      case 'É' : ret+= "%C3%88"
      break;
      case 'Í' : ret+= "%C3%8D"
      break;
      case 'Ó' : ret+= "%C3%93"
      break;
      case 'Ú' : ret+= "%C3%9A"
      break;
      case 'Ñ' : ret+= "%C3%91"
      break;        
      case 'ª' : ret+= "%2c%AA"
      break;
      case 'º' : ret+= "%c2%B0"		
      break;
      case "'" : ret+= "%27"
      break;
      case 'ü' : ret+= "%c3%bc"
      break;
      case 'Ü' : ret+= "%c3%9c"        
      break;
      default:
        ret += str[i]
    }
  }    
  return ret;
}

function generateVCard( name, title , tlf , email, addr, organization, url){
	var getvcard = "BEGIN%3AVCARD%0AVERSION%3A3.0";
	if(name) getvcard += "%0AN%3A"+name;
	if(organization) getvcard += "%0AORG%3A"+organization;
	if(title) getvcard += "%0ATITLE%3A"+title;
	if(tlf) getvcard += "%0ATEL%3A"+tlf;
	if(email) getvcard += "%0AEMAIL%3A"+email;
	if(addr) getvcard += "%0AADR%3A"+addr;
	if(url) getvcard += "%0AURL%3A"+url;
	getvcard += "%0AEND%3AVCARD";
	return getvcard;
}

function generateQRLink( rowIndex ){
  var range = sheet.getRange( rowIndex, 1, 1, 9);

  var col=2;
  var name = range.getCell(1,col++).getValue();
  var title = range.getCell(1,col++).getValue();
  var tlf=range.getCell(1,col++).getValue();
  var email=range.getCell(1,col++).getValue();
  var addr=range.getCell(1,col++).getValue();
  var organization=range.getCell(1,col++).getValue();
  var url=range.getCell(1,col++).getValue();
  
  var vcard = urlRoot + generateVCard(name, title , tlf , email, addr, organization, url)

  range.getCell(1,9).setValue(vcard); 
  return true;
}

function cleanQRLinks(){
  for(var rowIndex=2; rowIndex<=sheet.getLastRow();rowIndex++){
      var range = sheet.getRange( rowIndex, 1, 1, 9);
      range.getCell(1,9).setValue(''); 
  }
}

function populateQRLinks(group){
  cleanQRLinks();
  for(var i=2; i<=sheet.getLastRow();i++){
    if( ! generateQRLink(i) ){
      return;
    }
  }
  nextStage(4, group);
}


function _findFolder(group ){
  var folder;
  DriveApp.getRootFolder();
  var name = 'Contacts2QRCode_'+group;
  var folders = DriveApp.getFoldersByName(name);
  while( folders.hasNext() ){
    folder = folders.next();    
  }
  if( folder == null ){
    folder = DriveApp.createFolder(name)
  }
  return folder;
}

function saveQRLinks( group ){
  var folder = _findFolder(group);
  if( folder == null){
     var html = HtmlService.createTemplateFromFile('Error').evaluate()
               .setSandboxMode(HtmlService.SandboxMode.IFRAME);
      SpreadsheetApp.getUi().showModalDialog(html, 'Error');  
     return -1;
  }
  nextStage(5, group, folder);
}

function dumpRow(group, rowIndex){
    var folder = _findFolder(group);
    var range = sheet.getRange( rowIndex+1, 1, 1, 9);
    var id = range.getCell(1,1).getValue();
    var url = range.getCell(1,9).getValue();  
    var iter = folder.getFilesByName(""+id);    
    while(iter.hasNext()){        
      iter.next().setTrashed(true);
    }
    var file = folder.createFile( UrlFetchApp.fetch(url) );
    file.setName(id);
    return group;
}


function showFolder(group ){ 
  var folder = _findFolder(group);
  nextStage(6, group, folder);
}

function replaceIdWithName(){
   for(var i=2; i<=sheet.getLastRow();i++){
    sheet.getRange(i,1).setValue(sheet.getRange(i,2).getValue());
  } 
}

function replaceAddrWith(org){
   for(var i=2; i<=sheet.getLastRow();i++){
    sheet.getRange(i,6).setValue(org)
  } 
}

function replaceOrgWith(org){
   for(var i=2; i<=sheet.getLastRow();i++){
    sheet.getRange(i,7).setValue(org)
  } 
}

function replaceUrlWith(org){
   for(var i=2; i<=sheet.getLastRow();i++){
    sheet.getRange(i,8).setValue(org)
  } 
}
