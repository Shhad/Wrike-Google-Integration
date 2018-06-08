//In Google App Script I have script properties:
//				- "customFields" - my wrike custom fields id's which I want include in report
//				- "users" - my wrike users id's , I'm counting work hours which my users from specified group spent under the project

TOKEN: "YOURS WRIKE APPLICATION TOKEN";

//Function which creates custom menu
function onOpen() {
  SpreadsheetApp.getUi().createMenu("MENU NAME")
  .addItem("YOUR OPTION NAME", 'YOUR FUNCTION')
  .addSeparator()
  //... ADDING NEW ITEMS
  //My functions
  .addItem("Organize projects from wrike", 'organizeReport')
  .addSeparator()
  .addItem("Actualise wrike projects", 'createReport')
  .addSeparator()
  .addToUi();
}

//This function imports all wrike projects from specific folder
function organizeReport() {
  clear1();
  var sheet = SpreadsheetApp.getActive();
  var ss = sheet.getSheetByName("YOUR SHEET NAME");
  
  var wrikeAPP = "www.wrike.com/api/v3"; //check latest version on wrike api!
  //type of resources what we want
  var type = "/folders";
  //optionally if we want specific type of resources we can get it by id, for example id of folder
  var ID = "/YOUR ID";
  //our final url to resources
  var https = "https://" + wrikeAPP + type + ID;
  //store permanent token
  var permanent_token = TOKEN;
  // prepare params for request:
  var parameters = {
	
  };
  // Set up Authorization header using permanent token:
  parameters.headers = { Authorization: 'bearer ' + permanent_token};
  //make HTTP request
  //store response from server and send it to another function
  var resp = UrlFetchApp.fetch(https, parameters);
  var data = resp.getContentText();
  getFoldersO(data,ss);
}

//This function is actualising each project in spreadsheet
function createReport() {
  //get spreadsheet in which I create report
  var sheet = SpreadsheetApp.getActive();
  var ss = sheet.getSheetByName("PPZ");
  //get spreadsheet data
  var data = ss.getRange(1, 1, ss.getLastRow(), ss.getLastColumn()).getValues();
  //for each project
  for(var i = 1;i<ss.getLastRow();i++) {
    if(data[i][7].toString().equals("NIE")) {
	  //store all data which could be already written, because when wrike server will
	  //reject our request we need to know which data whas already taken from wrike database
      var fileds = new Array(4);
      fileds[0] = data[i][2];
      fileds[1] = data[i][3];
      fileds[2] = data[i][4];
      fileds[3] = data[i][5];
	  //try find all projects details which we want to import
      try {
        getFolders(data[i][6], (1+i), ss);
      } catch(e) {
		//if server will reject or some exception occurred
		//print error message and reset values in row
        Logger.log(e);
        ss.getRange((i + 1), 2).setValue(0);
        ss.getRange((i + 1), 3).setValue(fileds[0]);
        ss.getRange((i + 1), 4).setValue(fileds[1]);
        ss.getRange((i + 1), 5).setValue(fileds[2]);
        ss.getRange((i + 1), 6).setValue(fileds[3]);
        
        var error = new Array(1);
        error[0] = e.message;
		//Check if error occurred because wrike server rejected request
        if(error[0].length > 350) {
          ss.getRange((i + 1), 8).setValue("NIE");
          SpreadsheetApp.getUi().alert("Serwer Wrike odrzucił połączenie.\nNaciśnij AKTUALIZUJ jeszcze raz.");
          return;
        } else {
		  //if that was other error
          ss.getRange((i + 1), 8).setValue("BŁĄD KRYTYCZNY: " + e.message);
          Logger.log("Nastąpił błąd krytyczny podczas analizy zadania:" + ss.getRange((i + 1), 1).getValue() + "!!!!!!!!" );
        }
      }
    }
  }
}
  //clear4();
  var sheet = SpreadsheetApp.getActive();
  var ss = sheet.getSheetByName("ZAKOŃCZONE PPZ");
  
  var wrikeAPP = "www.wrike.com/api/v3";
  //type of resources what we want
  var type = "/folders";
  //optionally if we want specific type of resources we can get it by id
  var ID = "/IEABDEOYI4FFLUT5";
  //our final url to resources
  var https = "https://" + wrikeAPP + type + ID;
  //store permanent token
  var permanent_token = 'bpzlInhTtQHQPFaQGrgam3ILfDDmNuEwxEXwY2661JlX6iJAN0Y74JpDBrnG5Asc-N-WFIUKC';
  // prepare params for request:
  var parameters = {
    //agent: false
  };
  // Set up Authorization header using permanent token:
  parameters.headers = { Authorization: 'bearer ' + permanent_token};
  //make HTTP request
  //store response from server
  var resp = UrlFetchApp.fetch(https, parameters);
  var data = resp.getContentText();
  getFoldersO(data,ss);
}

//This function gets wrike projects from folder and goes through folder tree to find any projects in current folder
function getFoldersO(answer,ss) {
	//get title of project
  var title = getFolderTitle(answer);
  Logger.log(title);
  //I USED HERE RAW JSON PARSING
  //parsing answer in json format to get specified data(projects or folders in current folder)
  var IDivide = answer.split("[");
  
  var IIDivide = IDivide[4].split('"');

  var childsNum = parseInt((IIDivide.length - 3)/2);
  var childsID = new Array(childsNum);
  Logger.log("Number of childIDs : " + childsID.length);
  
  var childCounter = 0;
  //get each child ID(child ID is ID of folder
  for(var  i = 1 ;i < (IIDivide.length-2);) {
    childsID[childCounter] = IIDivide[i];
    childCounter++;
    i=i+2;
  }
  if(childCounter != 0) {
     //check if everything is good
    for(var  i = 0 ;i < childCounter;i++) {
      Logger.log(childsID[i]);
    }
    for(var  i = 0 ;i < childCounter;i++) {
      var wrikeAPP = "www.wrike.com/api/v3";
      var type = "/folders";
      var ID = "/" + childsID[i];
      var https = "https://" + wrikeAPP + type + ID;
      var permanent_token = TOKEN;
      var parameters = {
        
      };
      parameters.headers = { Authorization: 'bearer ' + permanent_token};
      
      var resp = UrlFetchApp.fetch(https, parameters);
      var data = resp.getContentText();
	  //Sleep for disable sending to many requests to API 
      Utilities.sleep(50);
      getFoldersO(data,ss);
    }
  } else {
	//If there are no folders inside
	//get project folder ID
	var folderID = getFolderID(answer);
	//check if project exist in spreadsheet
    var ppzZakonczone = checkFinished(folderID, ss);
	//if yes
    if(ppzZakonczone) {
	  //get next row in spreadsheet
      var row = checkProjectRow(title,ss);
	  //get custom fields
      IIDivide = IDivide[7].split('"');
      if(IIDivide[1].equals("project")) {
        Logger.log("No custom fields in project scope!");
      } else {
        Logger.log("There are custom fields in project scope!");
        for(var i = 3;i<IIDivide.length;) {
			//get custom fields
          var idc = IIDivide[i];
          i=i+4;
          var idv = IIDivide[i];
		  //write data from custom fields to spreadsheet
          fillCustomFieldF(idc,idv,row,ss);
          if(IIDivide[i+2].equals("project")) break;
          i = i + 4;
        }
      }
      //set in spreadsheet that project exist but it is not actual
      ss.getRange(row, 7).setValue(folderID);
      ss.getRange(row, 8).setValue("NIE").setBackground("#ff0000");
    }
	//if not, doing nothing
  }
}

//Function gets all tasks assigned to project
function getFolders(id, row, ss) {
  //same request for project tasks
  var wrikeAPP = "www.wrike.com/api/v3";
  var type = "/folders";
  var ID = "/" + id + "/tasks";
  var https = "https://" + wrikeAPP + type + ID;
  var permanent_token = TOKEN;
  var parameters = {

  };
  parameters.headers = { Authorization: 'bearer ' + permanent_token};
  var resp = UrlFetchApp.fetch(https, parameters);
  var data = resp.getContentText();
  getTasksFromFolder(data,row, ss);
}

//Function which clears sheet in spreadsheet
function clear1() {
  var sheet = SpreadsheetApp.getActive();
  var ss = sheet.getSheetByName("YOUR SHEET NAME");
  try {
    ss.deleteRows(2, (ss.getLastRow() - 1));
  } catch(e) {
    Logger.log(e);
  }
}
  var sheet = SpreadsheetApp.getActive();
  var ss = sheet.getSheetByName("ZAKOŃCZONE PPZ");
  try {
    ss.deleteRows(2, (ss.getLastRow() - 1));
  } catch(e) {
    Logger.log(e);
  }
}

//Function which checks if project exist in spreadsheet scope
function checkFinished(id, ss) {
  var data = ss.getRange(1, 1, ss.getLastRow(), ss.getLastColumn()).getValues();
  Logger.log("id: " + id);
  for(var  i = 1;i<ss.getLastRow();i++) {
    if(id.equals(data[i][6])) {
       return false;
    }
  }
  return true;
}

//Find next empty row
function checkProjectRow(title,ss) {
  var rowN = ss.getLastRow();
  ss.getRange(rowN+1, 1).setValue(title);
  return (rowN+1);
}

//Get folder title(project title) from server answer
function getFolderTitle(answer) {
  var IDivide = answer.split("[");
  //Logger.log(IDivide[1]);
  var IIDivide = IDivide[1].split('"');
  //Logger.log(IIDivide[11]);
  return IIDivide[11];
}

//Get folder ID(project ID) from server answer
function getFolderID(answer) {
  var IDivide = answer.split("[");
  //Logger.log(IDivide[1]);
  var IIDivide = IDivide[1].split('"');
  //Logger.log(IIDivide[3]);
  return IIDivide[3];
}

//Function is checking if in tasks scope are more folders(tasks might be hidden in subfolders)
function getTasksFromFolder(answer, row,ss) {
  
  var IDivide = answer.split('"id"');
  var tasksNum = IDivide.length - 1;
  var IIDivide;
  var taskID;
  for(var i = 1; i< IDivide.length;i++) {
	//The same request like in other functions
    IIDivide = IDivide[i].split('"');
    taskID = IIDivide[1];
    
    var wrikeAPP = "www.wrike.com/api/v3";
    var type = "/tasks";
    var ID = "/" + taskID;
    var https = "https://" + wrikeAPP + type + ID;
    var permanent_token = TOKEN;
    var parameters = {

    };
    parameters.headers = { Authorization: 'bearer ' + permanent_token};
	
    var resp = UrlFetchApp.fetch(https, parameters);
    var data = resp.getContentText();
    Utilities.sleep(50);
	//get task details
    getTasks(data, row, ss);
  }
  //get working hours of employees
  var duration = ss.getRange(row, 2).getValue();
  duration = parseInt(duration,10);
  if(duration==0) {
    ss.getRange(row, 2).setValue(duration);
  } else if(duration > 0) {
    duration = duration/60;
    //duration = duration/8;
    ss.getRange(row, 2).setValue(duration);
  } else {
    ss.getRange(row, 2).setValue(0);
  }
  
  ss.getRange(row, 8).setValue("TAK").setBackground("#00ff00");
  Logger.log("-------------------------------------------------------------");
}

//Function which gets task details and check if there are any subtasks
function getTasks(answer, row,ss) {
  Logger.log("------------------------------------------------------------------");
  //get task title
  var title = getFolderTitle(answer);
  Logger.log(title);
  //get subtasks
  var IDivide = answer.split("subTaskIds");
  var helpDivide = IDivide[1].split("[");
  var IIDivide = helpDivide[1].split('"');
  IDivide = answer.split("[");
  //Logger.log(IIDivide.length);
  var childsNum = parseInt((IIDivide.length - 3)/2);
  var childsID = new Array(childsNum);
  Logger.log("Numbers of childID : " + childsID.length);
  //Logger.log(IDivide[4]);
  var childCounter = 0;
  for(var  i = 1 ;i < (IIDivide.length-2);) {
    childsID[childCounter] = IIDivide[i];
    childCounter++;
    i=i+2;
  }
  
  if(childCounter != 0) {
     //check if everything is good
     for(var  i = 0 ;i < childCounter;i++) {
       Logger.log(childsID[i]);
     }
    var userID2 = checkUser(answer);
    if(userID2) {
      //if there are no subTasks we are searching for customFields
      try {
        IDivide = answer.split("customFields");
        //Logger.log(1);
        helpDivide = IDivide[1].split("[");
        //Logger.log(2);
        IIDivide = helpDivide[1].split('"');
        //Logger.log(3);
        IDivide = answer.split("responsibleIds");
        //Logger.log(4);
        var IIIDivide = IDivide[1].split("duration");
        //Logger.log(5);
        var dParts1 = IIIDivide[1].split(",");
        //Logger.log(6);
        var dParts2 = dParts1[0].split(" ");
        //Logger.log(7);
        var duration = dParts2[1];
        fillDuration(row,ss, duration);
      } catch(e) {
       Logger.log("no duration field!"); 
      }
      if(!(IIDivide.length <= 1)) {
        Logger.log("There are custom fields in tasks scope!");
        for(var i = 3;i<IIDivide.length;) {
          var idc = IIDivide[i];
          i=i+4;
          var idv = IIDivide[i];
          fillCustomFieldT(idc,idv,row, ss);
          if(i + 4 > IIDivide.length) break;
          i = i + 4;
        }
      }
    }
    for(var  i = 0 ;i < childCounter;i++) {
		//if task has any subtasks
      var wrikeAPP = "www.wrike.com/api/v3";
      var type = "/tasks";
      var ID = "/" + childsID[i];
      var https = "https://" + wrikeAPP + type + ID;
      var permanent_token = TOKEN;
      var parameters = {
        
      };
      parameters.headers = { Authorization: 'bearer ' + permanent_token};
      
      var resp = UrlFetchApp.fetch(https, parameters);
      var data = resp.getContentText();
	  //get details of subtask
      getTasks(data, row, ss);
    }
  } else {
    var userID = checkUser(answer);
    if(userID) {
      //if there are no subTasks we are searching for customFields
      try {
        IDivide = answer.split("customFields");
        //Logger.log(1);
        helpDivide = IDivide[1].split("[");
        //Logger.log(2);
        IIDivide = helpDivide[1].split('"');
        //Logger.log(3);
        IDivide = answer.split("responsibleIds");
        //Logger.log(4);
        var IIIDivide = IDivide[1].split("duration");
        //Logger.log(5);
        var dParts1 = IIIDivide[1].split(",");
        //Logger.log(6);
        var dParts2 = dParts1[0].split(" ");
        //Logger.log(7);
        var duration = dParts2[1];
        fillDuration(row,ss, duration);
      } catch(e) {
       Logger.log("no duration field!"); 
      }
      if(!(IIDivide.length <= 1)) {
        Logger.log("There are custom fields in tasks scope!");
        for(var i = 3;i<IIDivide.length;) {
          var idc = IIDivide[i];
          i=i+4;
          var idv = IIDivide[i];
          fillCustomFieldT(idc,idv,row, ss);
          if(i + 4 > IIDivide.length) break;
          i = i + 4;
        }
      }
    }
  }
}

//Function which fill spreadsheet columns which custom data fields taken from folder(project) scope
function fillCustomFieldF(CFID, value, row, ss) {
  var ids = PropertiesService.getScriptProperties().getProperty('customFields');
  ids = ids.split(",");
  for(var id = 0; id < ids.length;id++) {
    if(ids[id].equals(CFID)) {
      if(id == 0) {
        if(!value.equals("")) {
          ss.getRange(row, id + 3).setValue(parseInt(value,10));
          Logger.log("Wkladam dane do kolumny " + (id + 3) +" o wartosci: " + parseInt(value,10));
        }
      }
      if(id == 1) {
        if(!value.equals("")) {
          ss.getRange(row, id + 3).setValue(parseInt(value,10));
          Logger.log("Wkladam dane do kolumny " + (id + 3) +" o wartosci: " + parseInt(value,10));
        }
      }
      if(id == 2) {
        if(!value.equals("")) {
          ss.getRange(row, id + 3).setValue(parseInt(value,10));
          Logger.log("Wkladam dane do kolumny " + (id + 3) +" o wartosci: " + parseInt(value,10));
        }
      }
      if(id == 3) {
        if(!value.equals("")) {
          ss.getRange(row, id + 3).setValue(parseInt(value,10));
          Logger.log("Wkladam dane do kolumny " + (id + 3) +" o wartosci: " + parseInt(value,10));
        }
      }
    }
  }
}

//Function which fill spreadsheet columns which custom data fields taken from task scope
function fillCustomFieldT(CFID, value, row, ss) {
  var ids = PropertiesService.getScriptProperties().getProperty('customFields');
  ids = ids.split(",");
  var vc;
  for(var id = 0; id < ids.length;id++) {
    if(ids[id].equals(CFID)) {
      if(id == 0) {
        if(!value.equals("")) {
          vc = ss.getRange(row, id + 3).getValue();
          Logger.log("W komorce bylo juz: " + vc);
          if((vc.toString()).equals("")) {
            ss.getRange(row, id + 3).setValue(parseInt(value,10));
            Logger.log("Wkladam dane do kolumny " + (id + 3) +" o wartosci: " + parseInt(value,10));
          } else {
            vc = parseInt(vc,10) + parseInt(value,10);
            ss.getRange(row, id + 3).setValue(vc);
            Logger.log("Wkladam dane do kolumny " + (id + 3) +" o wartosci: " + vc);
          }
        }
      }
      if(id == 1) {
        if(!value.equals("")) {
          vc = ss.getRange(row, id + 3).getValue();
          Logger.log("W komorce bylo juz: " + vc);
          if((vc.toString()).equals("")) {
            ss.getRange(row, id + 3).setValue(parseInt(value,10));
            Logger.log("Wkladam dane do kolumny " + (id + 3) +" o wartosci: " + parseInt(value,10));
          } else {
            vc = parseInt(vc,10) + parseInt(value,10);
            ss.getRange(row, id + 3).setValue(vc);
            Logger.log("Wkladam dane do kolumny " + (id + 3) +" o wartosci: " + vc);
          }
        }
      }
      if(id == 2) {
        if(!value.equals("")) {
          vc = ss.getRange(row, id + 3).getValue();
          Logger.log("W komorce bylo juz: " + vc);
          if((vc.toString()).equals("")) {
            ss.getRange(row, id + 3).setValue(parseInt(value,10));
            Logger.log("Wkladam dane do kolumny " + (id + 3) +" o wartosci: " + parseInt(value,10));
          } else {
            vc = parseInt(vc,10) + parseInt(value,10);
            ss.getRange(row, id + 3).setValue(vc);
            Logger.log("Wkladam dane do kolumny " + (id + 3) +" o wartosci: " + vc);
          }
        }
      }
      if(id == 3) {
        if(!value.equals("")) {
         vc = ss.getRange(row, id + 3).getValue();
          Logger.log("W komorce bylo juz: " + vc);
          if((vc.toString()).equals("")) {
            ss.getRange(row, id + 3).setValue(parseInt(value,10));
            Logger.log("Wkladam dane do kolumny " + (id + 3) +" o wartosci: " + parseInt(value,10));
          } else {
            vc = parseInt(vc,10) + parseInt(value,10);
            ss.getRange(row, id + 3).setValue(vc);
            Logger.log("Wkladam dane do kolumny " + (id + 3) +" o wartosci: " + vc);
          }
        }
      }
    }
  }
}

//Function which fills duration column of project
function fillDuration(row, ss, duration) {
  var vc = ss.getRange(row, 2).getValue();
  var vcN = parseInt(vc,10);
  Logger.log("vc number: " + vcN);
  Logger.log("Duration: " + duration);
  var vcS = vc.toString();
  if(duration.equals("")) {
    duration = 0;
  }
  Logger.log("W komorce bylo juz: " + vc);
  Logger.log("Dl vc: " + vcS.length);
  if(vcS.length < 1) {
    ss.getRange(row, 2).setValue(parseInt(duration,10));
    Logger.log("Wkladam dane do kolumny " + (2) +" o wartosci: " + parseInt(duration,10));
  } else {
    vc = parseInt(vc,10) + parseInt(duration,10);
    ss.getRange(row, 2).setValue(vc);
    Logger.log("Wkladam dane do kolumny " + (2) +" o wartosci: " + vc);
  }
}

//Check is task assigned to user which we are searching for
function checkUser(answer) {
  var IDivide = answer.split("responsibleIds");
  var helpDivide = IDivide[1].split("[");
  var IIDivide = helpDivide[1].split('"');
  var id = IIDivide[1];
  if(id.length < 6) {
    return false;
  } else {
    var users = PropertiesService.getScriptProperties().getProperty('users').split(",");
    Logger.log("ID ktorego szukamy: " + id);
    for(var  i = 0;i<users.length;i++) {
      //Logger.log((i+1) + " ID: " + users[i]);
      if(users[i].equals(id)) {
        Logger.log("ZNaleziono szukanego uzytkownika!");
        return true;
      }
    }
  }
  return false;
}
  var ss = SpreadsheetApp.getActive().getSheetByName("PPZ");
  var row = 5;
  var id = "IEABDEOYI4GECCI3";
  //getFolders(id,row,ss);
  /*if(checkFinished(id,ss)) {
    Logger.log("nie ma");
  } else {
    Logger.log("jest");
  }*/
}
