function collect() {
  
  var satSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Input Sheet"); // identifies the sheets at the account level
  var eventSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Event List");
  var sd = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet Details");
  
 
  for(var i = 3; i < 253; i++){
    
    var event1 = satSheet.getRange(i,41).getValue(); // column 28 is where the label is 
    var event2 = satSheet.getRange(i,47).getValue(); // column 33 is where the second label is
    

    if(event2 !== "") { 
      
    var actRow = satSheet.getActiveCell().getRow(); //this is the row co-ordinate of the change to the Event 1 Category **not sure this is used anymore? and line below.
    eventSheet.getRange(1,4).setValue(satSheet.getActiveCell().getRow());//prints the row of the on Edited cell Q why do you need to define the var within the if clause not above it?
    
    
    var range = eventSheet.getRange(1,1).getValue(); // Defines the range of the events down to the last event currently on the list
   
    var newPrint = range+2; // Row co-ordinate of the new event
    var nP2 = newPrint+1; // Row below for the second layer
    var Contact = satSheet.getRange(i,2).getValue(); // Returns the name of the contact
    var user = sd.getRange(5,3).getValue(); // User's name
    var userComp = sd.getRange(6,3).getValue(); // User's company name
    // var userLink = satSheet.getRange(7,47).getValue(); // User's sheet link
    var userMail = sd.getRange(11,3).getValue(); // User's email
    var contactmail = satSheet.getRange(i,11).getValue(); // Returns the name of the contact's work email
    var contactpmail = satSheet.getRange(i,12).getValue(); // Returns the name of the contact's private email
    var eventNote1 = satSheet.getRange(i,40).getValue(); // Returns event note 1
    var eventLabel1 = satSheet.getRange(i,41).getValue(); // event label 1
    var eventURL1 = satSheet.getRange(i,42).getValue(); // event URL 1
    var eventBAU1 = satSheet.getRange(i,43).getValue(); // event BAU 1
    var eventIMP1 = satSheet.getRange(i,44).getValue(); // event IMP 1
    var eventDate1 = satSheet.getRange(i,45).getValue(); // event date 1
    var userNote = sd.getRange(12,3).getValue(); // Returns the user note
    var contactComp = satSheet.getRange(i,1).getValue(); // contact's company
    var contactLI = satSheet.getRange(i,30).getValue(); // contact's LinkedIn
    var contactRole = satSheet.getRange(i,13).getValue(); // contacts's role
    var eventNote2 = satSheet.getRange(i,46).getValue(); // Returns event note 2
    var eventLabel2 = satSheet.getRange(i,47).getValue(); // event label 2
    var eventURL2 = satSheet.getRange(i,48).getValue(); // event URL 2
    var eventBAU2 = satSheet.getRange(i,49).getValue(); // event BAU 2
    var eventIMP2 = satSheet.getRange(i,50).getValue(); // event IMP 2
    var eventDate2 = satSheet.getRange(i,51).getValue(); // event date 2
    var eventhstat = satSheet.getRange(i,16).getValue(); // event home static
      
      Logger.log(user);
    
    eventSheet.getRange(newPrint,5).setValue(i);
    eventSheet.getRange(newPrint,6).setValue(Contact);
    eventSheet.getRange(newPrint,7).setValue(contactComp); // Contact's company
    eventSheet.getRange(newPrint,8).setValue(contactLI); // Individual LinkedIn profile
    eventSheet.getRange(newPrint,9).setValue(contactmail);
    eventSheet.getRange(newPrint,10).setValue(contactpmail);
    eventSheet.getRange(newPrint,11).setValue(userMail);
    eventSheet.getRange(newPrint,2).setValue(user); // User name
    eventSheet.getRange(newPrint,3).setValue(userComp); // User's company
    // eventSheet.getRange(newPrint,4).setValue(userLink); // User's sheet link
    eventSheet.getRange(newPrint,12).setValue(userNote); // User note
    eventSheet.getRange(newPrint,13).setValue(contactRole); // Contact's role
    eventSheet.getRange(newPrint,14).setValue(eventNote1); // event 1 note 
    eventSheet.getRange(newPrint,15).setValue(eventLabel1); // event 1 label
    eventSheet.getRange(newPrint,16).setValue(eventURL1); // event 1 URL
    eventSheet.getRange(newPrint,17).setValue(eventDate1); // event 1 date
    eventSheet.getRange(newPrint,20).setValue(eventBAU1); // event 1 BAU
    eventSheet.getRange(newPrint,21).setValue(eventIMP1); // event 1 IMP
    eventSheet.getRange(newPrint,22).setValue(eventhstat); // event 1 home static
      
    eventSheet.getRange(nP2,5).setValue(i);
    eventSheet.getRange(nP2,6).setValue(Contact);
    eventSheet.getRange(nP2,7).setValue(contactComp); // Contact's company
    eventSheet.getRange(nP2,8).setValue(contactLI); // Individual LinkedIn profile
    eventSheet.getRange(nP2,9).setValue(contactmail);
    eventSheet.getRange(nP2,10).setValue(contactpmail);
    eventSheet.getRange(nP2,11).setValue(userMail);
    eventSheet.getRange(nP2,2).setValue(user); // User name
    eventSheet.getRange(nP2,3).setValue(userComp); // User's company
    // eventSheet.getRange(nP2,4).setValue(userLink); // User's sheet link
    eventSheet.getRange(nP2,12).setValue(userNote); // User note
    eventSheet.getRange(nP2,13).setValue(contactRole); // Contact's role
    eventSheet.getRange(nP2,14).setValue(eventNote2); // event 2 note 
    eventSheet.getRange(nP2,15).setValue(eventLabel2); // event 2 label
    eventSheet.getRange(nP2,16).setValue(eventURL2); // event 2 URL
    eventSheet.getRange(nP2,17).setValue(eventDate2); // event 2 date
    eventSheet.getRange(nP2,20).setValue(eventBAU2); // event 2 BAU
    eventSheet.getRange(nP2,21).setValue(eventIMP2); // event 2 IMP
    eventSheet.getRange(nP2,21).setValue(eventhstat); // event 2 home static
      
    } else if(event1 !== "") {
      
    var range = eventSheet.getRange(1,1).getValue(); // Defines the range of the events down to the last event currently on the list
   
    
    var newPrint = range+2; // Row co-ordinate of the new event
    var nP2 = newPrint+1; // Row below for the second layer
    var Contact = satSheet.getRange(i,2).getValue(); // Returns the name of the contact
    var user = sd.getRange(5,3).getValue(); // User's name
    var userComp = sd.getRange(6,3).getValue(); // User's company name
    // var userLink = satSheet.getRange(7,47).getValue(); // User's sheet link
    var userMail = sd.getRange(11,3).getValue(); // User's email
    var contactmail = satSheet.getRange(i,11).getValue(); // Returns the name of the contact's work email
    var contactpmail = satSheet.getRange(i,12).getValue(); // Returns the name of the contact's private email
    var eventNote1 = satSheet.getRange(i,40).getValue(); // Returns event note 1
    var eventLabel1 = satSheet.getRange(i,41).getValue(); // event label 1
    var eventURL1 = satSheet.getRange(i,42).getValue(); // event URL 1
    var eventBAU1 = satSheet.getRange(i,43).getValue(); // event BAU 1
    var eventIMP1 = satSheet.getRange(i,44).getValue(); // event IMP 1
    var eventDate1 = satSheet.getRange(i,45).getValue(); // event date 1
    var userNote = sd.getRange(12,3).getValue(); // Returns the user note
    var contactComp = satSheet.getRange(i,1).getValue(); // contact's company
    var contactLI = satSheet.getRange(i,30).getValue(); // contact's LinkedIn
    var contactRole = satSheet.getRange(i,13).getValue(); // contacts's role
    var eventhstat = satSheet.getRange(i,16).getValue(); // event home static  
    
    eventSheet.getRange(newPrint,5).setValue(i);
    eventSheet.getRange(newPrint,6).setValue(Contact);
    eventSheet.getRange(newPrint,7).setValue(contactComp); // Contact's company
    eventSheet.getRange(newPrint,8).setValue(contactLI); // Individual LinkedIn profile
    eventSheet.getRange(newPrint,9).setValue(contactmail);
    eventSheet.getRange(newPrint,10).setValue(contactpmail);
    eventSheet.getRange(newPrint,11).setValue(userMail);
    eventSheet.getRange(newPrint,2).setValue(user); // User name
    eventSheet.getRange(newPrint,3).setValue(userComp); // User's company
    // eventSheet.getRange(newPrint,4).setValue(userLink); // User's sheet link
    eventSheet.getRange(newPrint,12).setValue(userNote); // User note
    eventSheet.getRange(newPrint,13).setValue(contactRole); // Contact's role
    eventSheet.getRange(newPrint,14).setValue(eventNote1); // event 1 note 
    eventSheet.getRange(newPrint,15).setValue(eventLabel1); // event 1 label
    eventSheet.getRange(newPrint,16).setValue(eventURL1); // event 1 URL
    eventSheet.getRange(newPrint,17).setValue(eventDate1); // event 1 date
    eventSheet.getRange(newPrint,20).setValue(eventBAU1); // event 1 BAU
    eventSheet.getRange(newPrint,21).setValue(eventIMP1); // event 1 IMP  
    eventSheet.getRange(newPrint,22).setValue(eventhstat); // event 1 home static  
      
    }
  }
  
  
}
