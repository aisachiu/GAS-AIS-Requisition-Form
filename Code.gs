// onformSubmit - go through and create a Google Doc from template, replacing all text with text from the form

var requestedDocDescriptor = "Requisition Form"

function onFormSubmit() {
  var templateDocId = '1BFTeXl889EB6NMgBkkHfDTNi-o3aQVpTR5Mcla1da8k';
  //var destination = DriveApp.getFolderById('0B0B30i6AUCFlRklRbHN3TE5OaVU');
  var destination = DriveApp.getFolderById('1ZoTFBoC1ZOIfhU9l539u1gQelAu5EySp');
  
  var mySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses 1');
  var formData = mySheet.getDataRange().getValues();
  
  //Get Col values
  var gDocCol = formData[0].indexOf('GDoc'); //Get col num oif the column we are writing the GDoc in
  var usernameCol = formData[0].indexOf('Email Address');
  var lineCost1Col = formData[0].indexOf('Line Cost');
  var lineCost2Col = formData[0].indexOf('Line Cost 2');
  var totalCostCol = formData[0].indexOf('Total Cost');
  
  //Loop through and fill in formulas first
  for (var i = 1; i < formData.length; i++){
    if (formData[i][gDocCol] === ""){ //if no GDoc has been created yet
      var row = i+1;
      mySheet.getRange(i+1, lineCost1Col+1, 1, 3).setValues([['=F'+row+'*G'+row, '=M'+row+'*N'+row, '=S'+row+'+R'+row]]);
      var totalCost, lineCost1, lineCost2 = 0;
      }
  }
  
  formData = mySheet.getDataRange().getValues();
  
  for (var i = 1; i < formData.length; i++){
    if (formData[i][gDocCol] === ""){ //if no GDoc has been created yet
      var row = i+1;
      mySheet.getRange(i+1, gDocCol+1).setValue("-1");
      var totalCost, lineCost1, lineCost2 = 0;
      /*try {
        lineCost1 = formData[i][5] * formData[i][6]
        mySheet.getRange(i+1, lineCost1Col).setValue(lineCost1);
        totalCost += lineCost1; 
        lineCost2 = formData[i][12] * formData[i][13]
        mySheet.getRange(i+1, lineCost2Col).setValue(lineCost2);
        totalCost += lineCost2;
        mySheet.getRange(i+1, totalCostCol).setValue(totalCost);
      } catch (e){}
*/
      var nowDate = Utilities.formatDate(new Date(), "GMT+08:00", "yyyy-MM-dd HH:mm");
      var filename = "New Requisition - " + formData[i][usernameCol] + " - " + nowDate;
      var newDocId = DriveApp.getFileById(templateDocId).makeCopy(filename, destination).getId();
      var newDoc = DocumentApp.openById(newDocId);
      var myText = newDoc.getBody();
      
      for (var cols = 0; cols < formData[0].length; cols++){
        var searchTxt = "{{"+formData[0][cols]+"}}";
        var replaceTxt = formData[i][cols];
        myText.replaceText(searchTxt, replaceTxt);    
      }//end for cols
      
      
      
      mySheet.getRange(i+1, gDocCol+1).setValue(newDocId);
      DriveApp.getFileById(newDoc.getId()).addEditor(formData[i][usernameCol]);
      
      var bodyText = 'You requested a ' + requestedDocDescriptor+'. Here it is: https://docs.google.com/document/d/' + newDocId;
      MailApp.sendEmail(formData[i][usernameCol], requestedDocDescriptor + ' Request', bodyText);

    } //end if GDoc empty
  }//end for every formData
}
