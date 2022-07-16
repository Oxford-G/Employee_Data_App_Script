// Form to validate entry made by user in User Form

function validateEntry() {
  // declare a variable and set the reference of active google sheet

  var myGoogleSheet=SpreadsheetApp.getActiveSpreadsheet();

  var shUserForm=myGoogleSheet.getSheetByName("User Form");

  var ui=SpreadsheetApp.getUi(); // to create the instance of the user interface to show the alert

  shUserForm.getRange("C7").setBackground('#ffffff');
  shUserForm.getRange("C9").setBackground('#ffffff');
  shUserForm.getRange("C11").setBackground('#ffffff');
  shUserForm.getRange("C13").setBackground('#ffffff');
  shUserForm.getRange("C15").setBackground('#ffffff');
  shUserForm.getRange("C17").setBackground('#ffffff');

  //Validating Employee ID
  if(shUserForm.getRange("C7").isBlank()==true) {
    ui.alert("Please enter Employee ID");
    shUserForm.getRange("C7").activate();
    shUserForm.getRange("C7").setBackground('#ff0000');
    return false;
  }

  //Validating Employee Name
  if(shUserForm.getRange("C9").isBlank()==true) {
    ui.alert("Please enter Employee Name");
    shUserForm.getRange("C9").activate();
    shUserForm.getRange("C9").setBackground('#ff0000');
    return false;
  }

  //Validating Gender
  if(shUserForm.getRange("C11").isBlank()==true) {
    ui.alert("Please select Gender from the drop-down");
    shUserForm.getRange("C11").activate();
    shUserForm.getRange("C11").setBackground('#ff0000');
    return false;
  }

  //Validating Email ID
  if(shUserForm.getRange("C13").isBlank()==true) {
    ui.alert("Please enter valid Email ID");
    shUserForm.getRange("C13").activate();
    shUserForm.getRange("C13").setBackground('#ff0000');
    return false;
  }

  //Validating Department
  if(shUserForm.getRange("C15").isBlank()==true) {
    ui.alert("Please select Department name from the drop-down");
    shUserForm.getRange("C15").activate();
    shUserForm.getRange("C15").setBackground('#ff0000');
    return false;
  }

  //Validating Address
  if(shUserForm.getRange("C17").isBlank()==true) {
    ui.alert("Please enter valid Address");
    shUserForm.getRange("C17").activate();
    shUserForm.getRange("C17").setBackground('#ff0000');
    return false;
  }

  return true;
}


//function to submit data to database sheet

function submitData(){
  
  //declare variable and set reference of active google sheet

  var myGoogleSheet=SpreadsheetApp.getActiveSpreadsheet();

  var shUserForm=myGoogleSheet.getSheetByName("User Form");

  var dataSheet=myGoogleSheet.getSheetByName("Database");

  //to create the instance of the user-interface enironment to use the alert feature

  var ui=SpreadsheetApp.getUi();

  var response=ui.alert("Submit", "Do you want to submit the data", ui.ButtonSet.YES_NO);

  //checking the user response

  if(response==ui.Button.NO){
    return // exit
  }

  if(validateEntry()==true){
    var blankRow=dataSheet.getLastRow()+1; // identify the next blank row

    // code to update the data in database sheet

    dataSheet.getRange(blankRow,1).setValue(shUserForm.getRange("C7").getValue()); // Employee ID

    dataSheet.getRange(blankRow,2).setValue(shUserForm.getRange("C9").getValue()); // Employee Name

    dataSheet.getRange(blankRow,3).setValue(shUserForm.getRange("C11").getValue()); // Gender

    dataSheet.getRange(blankRow,4).setValue(shUserForm.getRange("C13").getValue()); // Email ID

    dataSheet.getRange(blankRow,5).setValue(shUserForm.getRange("C15").getValue()); // Department

    dataSheet.getRange(blankRow,6).setValue(shUserForm.getRange("C17").getValue()); // Address

    //code to update the date and time -Submitted On

    dataSheet.getRange(blankRow,7).setValue(new Date()).setNumberFormat('yyyy-mm-dd h:mm');

    //Submitted by

    dataSheet.getRange(blankRow,8).setValue(Session.getActiveUser().getEmail());

    ui.alert('"New Date Saved - Emp # ' + shUserForm.getRange("C7").getValue() + '"');

      shUserForm.getRange("C7").clear();
      shUserForm.getRange("C9").clear();
      shUserForm.getRange("C11").clear();
      shUserForm.getRange("C13").clear();
      shUserForm.getRange("C15").clear();
      shUserForm.getRange("C17").clear();
  }
}

// function to serach the record

function searchRecord() {

  // declare variable and set with active google sheet

  var myGoogleSheet=SpreadsheetApp.getActiveSpreadsheet();

  //declare variable and set with user form reference

  var shUserForm=myGoogleSheet.getSheetByName("User Form");

  // declare variable and set the reference of Database sheet

  var dataSheet=myGoogleSheet.getSheetByName("Database");

  var str=shUserForm.getRange("C4").getValue();

  // getting the entire values from the used range and assigning it to values variable

  var values= dataSheet.getDataRange().getValues();

  var valuesFound= false; // variable to store boolean values

  for (i=0; i<values.length; i++){

    var rowValue= values[i];

    // checking the first value of the record is equal to search item

    if (rowValue[0]==str) {
      shUserForm.getRange("C7").setValue(rowValue[0]);
      shUserForm.getRange("C9").setValue(rowValue[1]);
      shUserForm.getRange("C11").setValue(rowValue[2]);
      shUserForm.getRange("C13").setValue(rowValue[3]);
      shUserForm.getRange("C15").setValue(rowValue[4]);
      shUserForm.getRange("C17").setValue(rowValue[5]);
      valuesFound= true;
      return; //exit the loop
    }
  }

  if (valuesFound=false) {
    // to create the instance of the user-interface environment to use the alert function
    var ui= SpreadsheetApp.getUi();
    ui.alert("No record found!");
  }

}















