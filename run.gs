function run() {

  var ss = SpreadsheetApp.getActiveSpreadsheet(); //get active sheet doc
  var sheet = ss.getSheetByName("Sources + Destinations"); //get sheet of possible copy actions
  var lastRow = sheet.getLastRow()+1; //get last row of sheet
  var list = sheet.getRange("A1:AH"+lastRow).getValues(); //get all values in list of possible copy actions
  
  for (var z = 0; z < list.length; z++) {
    if (list[z][0]==true) { //if action is marked as active...
      write(list[z][1], //calls the write function which takes all the arguments and then writes the data
      list[z][3],
      list[z][4],
      list[z][5],
      list[z][6],
      list[z][7],
      list[z][8],
      list[z][9],
      list[z][10],
      list[z][11],
      list[z][12],
      list[z][13],
      list[z][14],
      list[z][15],
      list[z][16],      
      list[z][17],
      list[z][18],
      list[z][19],
      list[z][20],
      list[z][21]);
    }
    if (list[z][22]==true) { //if "roster" is true...
      roster(list[z][5],  //calls the roster function which compares two columns of student numbers and adds missing items
      list[z][8],
      list[z][9],
      list[z][23],
      list[z][24],
      list[z][25]);      
    }
  }

}


//function to write data using several inputs
function write(name, sourceKey, sourceTab, header, sourceCol1, sourceCol2, 
              destKey, destTab, destCol1, destRow, destCol2, 
              critCol1, crit1, critCol2, crit2, critCol3, crit3, tf1, tf2, tf3) {
                   
  var source = SpreadsheetApp.openById(sourceKey);
  var sourceTab = source.getSheetByName(sourceTab);
  var lastRow = sourceTab.getLastRow() + 1;
  var rawData = sourceTab.getRange(sourceCol1 + ":" + sourceCol2).getValues();
  
  var dest = SpreadsheetApp.openById(destKey);
  var destTab = dest.getSheetByName(destTab);
  var cleanData = []

  //add headers
  for (var x = 0; x < header; x++) {
    cleanData.push(rawData[x]);
  }
  
  //filter data when three filters exist
  if (critCol1 != "" && critCol2 != "" && critCol3 != "") {
  Logger.log("3");
     for (var i = header + 1; i < rawData.length; i++) {
        if(tf1 && rawData[i][critCol1] == crit1) {
          if(tf2 && rawData[i][critCol2] == crit2) {
            if(tf3 && rawData[i][critCol3] == crit3) {
              cleanData.push(rawData[i]);
            }
            else if(tf3 == false && rawData[i][critCol3] != crit3) {
              cleanData.push(rawData[i]);
            }
          }
          else if(tf2 == false && rawData[i][critCol2] != crit2) {
            if(tf3 && rawData[i][critCol3] == crit3) {
              cleanData.push(rawData[i]);
            }
            else if(tf3 == false && rawData[i][critCol3] != crit3) {
              cleanData.push(rawData[i]);
            }
          }
        }
        else if(tf1 == false && rawData[i][critCol1] != crit1) {
          if(tf2 && rawData[i][critCol2] == crit2) {
            if(tf3 && rawData[i][critCol3] == crit3) {
              cleanData.push(rawData[i]);
            }
            else if(tf3 == false && rawData[i][critCol3] != crit3) {
              cleanData.push(rawData[i]);
            }
          }
          else if(tf2 == false && rawData[i][critCol2] != crit2) {
            if(tf3 && rawData[i][critCol3] == crit3) {
              cleanData.push(rawData[i]);
            }
            else if(tf3 == false && rawData[i][critCol3] != crit3) {
              cleanData.push(rawData[i]);
            }
          }
        }    
     }
  }
  
  //filter data when two filters exist
  if (critCol1 != "" && critCol2 != "" && critCol3 == "") {
  Logger.log("2");
     for (var i = header + 1; i < rawData.length; i++) {
       if(tf1 == true && rawData[i][critCol1] == crit1) {
         if(tf2 == true && rawData[i][critCol2] == crit2) {
           cleanData.push(rawData[i]);
         }
         else if(tf2 == false && rawData[i][critCol2] != crit2) {
           cleanData.push(rawData[i]);
         }
       }
       
       else if (tf1 == false && rawData[i][critCol1] != crit1) {
         if(tf2 == true && rawData[i][critCol2] == crit2) {
           cleanData.push(rawData[i]);
         }
         else if(tf2 == false && rawData[i][critCol2] != crit2) {
           cleanData.push(rawData[i]);
         }       
       }
    }
  }
  
  //filter for data when only one filter exists
  if (critCol1 != "" && critCol2 == "" && critCol3 == "") {
  Logger.log("1");
     for (var i = header + 1; i < rawData.length; i++) {
       if(tf1 == true && rawData[i][critCol1] == crit1) {
         cleanData.push(rawData[i]);
       }
       else if(tf1 == false && rawData[i][critCol1] != crit1) {
         cleanData.push(rawData[i]);
       }
     }
  }

  //filter for data when only no filter exists
  if (critCol1 == "" && critCol2 == "" && critCol3 == "") {
      cleanData = rawData;
  }
  
  destTab.getRange(destCol1 + destRow + ":" + destCol2).clearContent(); //clear destination tab of data
  destTab.getRange(destCol1 + destRow + ":" + destCol2 + (cleanData.length - 1 + destRow)).setValues(cleanData); //set values
  
  log(name, cleanData); //log writing event
  
}

//helper function to log writing of data
function log (name, cleanData) {
  var log = SpreadsheetApp.getActive().getSheetByName("Log");
  var row = log.getLastRow()+1;
  var d = new Date();
  var curDate = Utilities.formatDate(d, "GMT-6", "MM/dd/yyyy");
  var curTime = d.toLocaleTimeString();
  log.getRange(row,1).setValue(curDate);
  log.getRange(row,2).setValue(curTime);
  log.getRange(row,3).setValue(name);
  log.getRange(row,4).setValue(cleanData.length);
}

function roster(headerRow, sourceKey, sourceTab, destTab, sourceCol, destCol) {
  var source = SpreadsheetApp.openById(sourceKey);
  var sourceTab = source.getSheetByName(sourceTab);
  var lastRow = sourceTab.getLastRow() + 1;
  headerRow++
  var twoD_rosterA = sourceTab.getRange(sourceCol + headerRow + ":" + sourceCol + lastRow).getValues();
  var destTab = source.getSheetByName(destTab);
  var lastRow2 = destTab.getLastRow();
  var twoD_rosterB = destTab.getRange(destCol + "1:" + destCol + lastRow2).getValues();
  var count = twoD_rosterB.filter(String).length; //find number of values in rosterB
  var blanks = [];
  count++;
  
  var rosterA = [];
  var rosterB = [];
  
  //flatten 2d roster A
  for (var line in twoD_rosterA) {
    rosterA.push(twoD_rosterA[line][0]);
  }
  
  //flatten 2d roster B
  for (var line in twoD_rosterB) {
    if(twoD_rosterB[line][0] != '') {
      rosterB.push(twoD_rosterB[line][0]);
    }
  }
    
  //compare roster A to roster B and add missing values to bottom of roster B
  var sort = false;
  for(var line in rosterA) {
    if(rosterB.indexOf(rosterA[line]) == -1) {
      if(sort === false) {
        destTab.sort(1);
        sort = true;
      }
      destTab.getRange(destCol+count).setValue(rosterA[line]);
      count++;
    }
  }

}
