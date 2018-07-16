// SHEET IDS

// Robustified Index Master: 1sEjzhq96me6aaQLIBqY6Wfgy9D6VrKhtHL9eUoqyT2Q
// Test Rig iQA: 1-5Vf4LbGOI29eVabBluk8WoHg5-8qJkzhgazLdLtVDE
// 9001 NE and U - Non-Events and Unchanged: 1-R-rShtcmrvJ1UKjEC-LZ2lungaLGtRW0dC3DwLXf8I
// 1201 Analysts Notes: 1q_gMJRlJzEgdnlN0Zx2LtSIurEpokrroteUH87ixmAc
// iDatabase: 1PDSr53kxFwWGDk9CdEu3KGu8mysSMxtExj7VRXV13nY
// 1301 Data Storage: 1W8ECF6uqytFJJ927CH3Z5-Ki5sYR0mgv69UWHRt-wSk


function collect() {
  // creates 3 arrays of any existing events. ret remains, QA goest to Events Queue and arch to 9001 NE and U archive
  var analyst = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet Details").getRange(7,3).getValue(); // finds the email address of the analyst on the Sheet Details
  var tDate = new Date();
  var ss3 = SpreadsheetApp.openById("1W8ECF6uqytFJJ927CH3Z5-Ki5sYR0mgv69UWHRt-wSk").getSheetByName("Sheet1"); // 1301 Data Storage
  var SQL = ss3.getRange("A1:A2000").getValues();
  var arcss1 = SpreadsheetApp.openById("1-R-rShtcmrvJ1UKjEC-LZ2lungaLGtRW0dC3DwLXf8I").getSheetByName("Sheet1"); // 9001 NE and U - Non-Events and Unchanged
   var ss1 = SpreadsheetApp.openById("1q_gMJRlJzEgdnlN0Zx2LtSIurEpokrroteUH87ixmAc").getSheetByName("Sheet1"); // 1201 Analysts Notes
  var l3 = ss1.getLastRow();
  var arcss2 = arcss1.getLastRow();
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("New Events");
  var l2 = ss.getLastRow();
  var ssEL = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Event List");
  var p2 = ssEL.getRange(1,1).getValue();
  var prep = ss.getRange("k2:k").getValues(); // column of events as per their id numbers
  var last = prep.filter(String).length; // number of events
  var data = ss.getRange(2,1,last,36).getValues(); // array of events
  var QA = new Array();
  var ret = new Array();
  var arch = new Array();
  
  var notes = ss.getRange(2,1,l2,35).getNotes();
  var beef = ss.getRange(2,1,l2,35).getValues();
  var notey = new Array();
  
  // the section below sweeps the notes from the analysts sheets to 1201 Analysts Notes
  for(z in notes){
    var row = notes[z];
   var bino = row.join();
    var bi = bino.length;
    if(bi >34){ 
     var blah = beef[z];
      notey.push(blah);
      notey.push(row);
    }
  }
  if(notey.length >0){
   ss1.getRange(l3+1, 1, notey.length, notey[0].length).setValues(notey);  
   ss.clearNotes();
  }
  
  // this next part distributes the events between 9001, 'Event List' and remaining +++ pushings NE events to 1301 Data Storage
  
  for(y in data){  // this clause creates a new array of unlabelled events and prints them back onto the sheet
    
    var label = data[y][5];  // the label column
    var labelled = true;
    var row = data[y];
    if(label == ""){  
      ret.push(row);
      var labelled = false;
    }
    
    if(label == "NO CHANGE" || label == "NON-EVENT"){  // this clause pushes NC/NE events to 9001 && 1301
      arch.push(row);
      for (var z = 0; z<2000; z++){
        var blah = SQL[z];
        var lame = blah.toString();
        if(lame.indexOf("indexed1")<0 && lame.indexOf(data[y][10])>=0 &&  lame.indexOf(data[y][0])>=0 && lame.indexOf(data[y][2])>=0){ // lame.indexOf("indexed1")<0 && 
            var big = [blah,analyst,tDate,label,"indexed1"];
            var newblah = big.join();
            ss3.getRange(z+1,1).setValue(newblah);
          
        }      
      }
     
      var labelled = false;
    }
    if(labelled == true){
       QA.push(row);
    } 
  }
  
 ss.getRange(2,1,l2,37).clearContent();
    
    if(ret.length >0){
 ss.getRange(2, 1, ret.length, ret[0].length).setValues(ret);
    }
  
    
    if(arch.length >0){
 arcss1.getRange(arcss2+1, 1, arch.length, arch[0].length).setValues(arch);
    }
    
    if(QA.length >0){

  var p3 = p2+2;
  for(var m = 0; m < QA.length; m++){
  ssEL.getRange(p3,2).setValue(QA[m][0]);
  ssEL.getRange(p3,3).setValue(QA[m][28]); // Company
    
  ssEL.getRange(p3,6).setValue(QA[m][2]);
  ssEL.getRange(p3,7).setValue(QA[m][1]); // Account
  ssEL.getRange(p3,8).setValue(QA[m][12]); // LinkedIn
  ssEL.getRange(p3,9).setValue(QA[m][29]); // CONtact email ----
  ssEL.getRange(p3,11).setValue(QA[m][30]); // user email
  ssEL.getRange(p3,12).setValue(QA[m][32]); // user story
  ssEL.getRange(p3,13).setValue(QA[m][31]); // contact role
  ssEL.getRange(p3,14).setValue(QA[m][6]); // Event note
  ssEL.getRange(p3,15).setValue(QA[m][5]); // Event Label
  ssEL.getRange(p3,16).setValue(QA[m][4]); // event URL
  ssEL.getRange(p3,17).setValue(QA[m][9]); // event date
  ssEL.getRange(p3,20).setValue(QA[m][7]); // BAU
  ssEL.getRange(p3,21).setValue(QA[m][8]); // IMPACT
  ssEL.getRange(p3,22).setValue(QA[m][11]); // Home page
  ssEL.getRange(p3,23).setValue(QA[m][19]); // last bmail
  ssEL.getRange(p3,24).setValue(QA[m][10]); // ID
  
    
  var p3 = p3+1;
  }
      
    
}
}

  
function oldcollect() {
  
  var satSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Input Sheet"); // identifies the sheets at the account level
  var eventSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Event List");
  var sd = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet Details");
  
 
  for(var i = 3; i < 253; i++){
    
    var event1 = satSheet.getRange(i,41).getValue(); // column 41 is where the label is 
    var event2 = satSheet.getRange(i,47).getValue(); // column 47 is where the second label is
    

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
    var eventlbmail = satSheet.getRange(i,54).getValue(); // last bmail for that contact 
      
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
    eventSheet.getRange(newPrint,23).setValue(eventlbmail); // event bmail
      
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
    eventSheet.getRange(nP2,22).setValue(eventhstat); // event 2 home static
    eventSheet.getRange(nP2,23).setValue(eventlbmail); // event bmail
    
      
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
    var eventlbmail = satSheet.getRange(i,54).getValue(); // last bmail for that contact 
    
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
    eventSheet.getRange(newPrint,23).setValue(eventlbmail); // last bmail
      
    }
  }
  satSheet.getRange("an3:ba300").clearContent();
    
  
}
