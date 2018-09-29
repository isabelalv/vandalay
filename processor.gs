function driver(){
  getData();
  processData();
  makeSlide();
}

function getData() {
  var url = 'https://my.api.mockaroo.com/isabel-large-response.json?key=e6ac1da0';
  //calling the API
  var response = UrlFetchApp.fetch(url, {'muteHttpExceptions': true});
  var arrayResponse = JSON.parse(response.getContentText());
  
  Logger.log(arrayResponse);
  //preparing the data to be put in the spreadsheet
  var dataarray = new Array();
  for (var i = 0; i<arrayResponse.length; i++){
    var thisRow = new Array();
    thisRow[0] = arrayResponse[i].id;
    thisRow[1] = arrayResponse[i].import_country; 
    thisRow[2] = arrayResponse[i].model;
    thisRow[3] = arrayResponse[i].make;
    thisRow[4] = arrayResponse[i].sold_by;
    thisRow[5] = arrayResponse[i].sale_price;
    
    dataarray[i] = thisRow;
  }
  
  //putting data in the spreadsheet
  var datasheet = SpreadsheetApp.openById("1KFbVpqHaUv3EsNaFXK9dpm_hiKs835Wp_mt57YZvlu0").getSheetByName("Data");
  datasheet.clear();
  datasheet.getRange(1, 1).setValue("ID");
  datasheet.getRange(1, 2).setValue("Import Country");
  datasheet.getRange(1, 3).setValue("Model");
  datasheet.getRange(1, 4).setValue("Make");
  datasheet.getRange(1, 5).setValue("Sold By");
  datasheet.getRange(1, 6).setValue("Sale Price");
  
  for(var i = 0; i<dataarray.length; i++){
    datasheet.getRange(i+2, 1).setValue(String(dataarray[i][0]));
    datasheet.getRange(i+2, 2).setValue(String(dataarray[i][1]));
    datasheet.getRange(i+2, 3).setValue(String(dataarray[i][2]));
    datasheet.getRange(i+2, 4).setValue(String(dataarray[i][3]));
    datasheet.getRange(i+2, 5).setValue(String(dataarray[i][4]));
    datasheet.getRange(i+2, 6).setValue(String(dataarray[i][5]));
  }
}

function processData(){
  processCountries();
  sortModels();
  sortMakes();
  sortSellers();
  leastTargetedCountries();
}

function processCountries(){
  var hashMap = new Object();
  var sales = new Object();
  var norepeat = {};
  var times = {};
  var countries = SpreadsheetApp.openById("1KFbVpqHaUv3EsNaFXK9dpm_hiKs835Wp_mt57YZvlu0").getSheetByName("Data").getRange(2, 2, 1001, 1).getValues();
  var salesData = SpreadsheetApp.openById("1KFbVpqHaUv3EsNaFXK9dpm_hiKs835Wp_mt57YZvlu0").getSheetByName("Data").getRange(2, 6, 1001, 1).getValues();
  var count = 0;
  
  //initializing hash maps
  for (var i = 0; i < countries.length; i++){
    hashMap[countries[i].toString()] = 0;
    sales[countries[i].toString()] = 0;
  }
  
  //counting occurrences of countries
  for (var i = 0; i < countries.length; i++){
    if(hashMap[countries[i].toString()] == 0){
      norepeat[count] = countries[i].toString();
      count ++;
    }
    hashMap[countries[i].toString()]++;
    sales[countries[i].toString()] += parseInt(salesData[i]);
    
    
  }
  
  var sheet = SpreadsheetApp.openById("1KFbVpqHaUv3EsNaFXK9dpm_hiKs835Wp_mt57YZvlu0").getSheetByName("Processed");
  for (var k = 0; k < count-1; k++){
    sheet.getRange(k+2, 7).setValue(norepeat[k]);
    sheet.getRange(k+2, 8).setValue(hashMap[norepeat[k]]);
    sheet.getRange(k+2, 9).setValue(sales[norepeat[k]]);
  }
  
  sheet.getRange(2, 7, 1000, 3).sort({column: 9, ascending: false});
}

function sortModels(){
  var hashMap = new Object();
  var norepeat = {};
  var times = {};
  var models = SpreadsheetApp.openById("1KFbVpqHaUv3EsNaFXK9dpm_hiKs835Wp_mt57YZvlu0").getSheetByName("Data").getRange(2, 3, 1001, 1).getValues();
  var count = 0;
  
  //initializing hash map
  for (var i = 0; i < models.length; i++){
    hashMap[models[i].toString()] = 0;
  }
  
  //counting occurrences of models
  for (var i = 0; i < models.length; i++){
    if(models[i].toString() != "undefined"){
      if(hashMap[models[i].toString()] == 0){
        norepeat[count] = models[i].toString();
        count ++;
      }
      hashMap[models[i].toString()]++;
    }
  }
  
  var sheet = SpreadsheetApp.openById("1KFbVpqHaUv3EsNaFXK9dpm_hiKs835Wp_mt57YZvlu0").getSheetByName("Processed");
  for (var k = 0; k < count-1; k++){
    sheet.getRange(k+2, 1).setValue(norepeat[k]);
    sheet.getRange(k+2, 2).setValue(hashMap[norepeat[k]]);
  }
  
  sheet.getRange(2, 1, 1000, 2).sort({column: 2, ascending: false});
}

function sortMakes(){
  var hashMap = new Object();
  var norepeat = {};
  var times = {};
  var makes = SpreadsheetApp.openById("1KFbVpqHaUv3EsNaFXK9dpm_hiKs835Wp_mt57YZvlu0").getSheetByName("Data").getRange(2, 4, 1001, 1).getValues();
  var count = 0;
  
  //initializing hash map
  for (var i = 0; i < makes.length; i++){
    hashMap[makes[i].toString()] = 0;
  }
  
  //counting occurrences of makes
  for (var i = 0; i < makes.length; i++){
    if(makes[i].toString() != "undefined"){
      if(hashMap[makes[i].toString()] == 0){
        norepeat[count] = makes[i].toString();
        count ++;
      }
      hashMap[makes[i].toString()]++;
    }
  }
  
  var sheet = SpreadsheetApp.openById("1KFbVpqHaUv3EsNaFXK9dpm_hiKs835Wp_mt57YZvlu0").getSheetByName("Processed");
  for (var k = 0; k < count-1; k++){
    sheet.getRange(k+2, 4).setValue(norepeat[k]);
    sheet.getRange(k+2, 5).setValue(hashMap[norepeat[k]]);
  }
  
  sheet.getRange(2, 4, 1000, 2).sort({column: 5, ascending: false});
}

function sortSellers(){
  var hashMap = new Object();
  var norepeat = {};
  var times = {};
  var sellers = SpreadsheetApp.openById("1KFbVpqHaUv3EsNaFXK9dpm_hiKs835Wp_mt57YZvlu0").getSheetByName("Data").getRange(2, 5, 1001, 1).getValues();
  var salesData = SpreadsheetApp.openById("1KFbVpqHaUv3EsNaFXK9dpm_hiKs835Wp_mt57YZvlu0").getSheetByName("Data").getRange(2, 6, 1001, 1).getValues();
  var count = 0;
  
  //initializing hash map
  for (var i = 0; i < sellers.length; i++){
    hashMap[sellers[i].toString()] = 0;
  }
  
  //counting occurrences of countries
  for (var i = 0; i < sellers.length; i++){
    if(hashMap[sellers[i].toString()] == 0){
      norepeat[count] = sellers[i].toString();
      count ++;
    }
    hashMap[sellers[i].toString()] += parseInt(salesData[i]);
  }
  
  var sheet = SpreadsheetApp.openById("1KFbVpqHaUv3EsNaFXK9dpm_hiKs835Wp_mt57YZvlu0").getSheetByName("Processed");
  for (var k = 0; k < count-1; k++){
    sheet.getRange(k+2, 11).setValue(norepeat[k]);
    sheet.getRange(k+2, 12).setValue(hashMap[norepeat[k]]);
  }
  
  sheet.getRange(2, 11, 1000, 2).sort({column: 12, ascending: false});
}

function leastTargetedCountries(){
  var sheet = SpreadsheetApp.openById("1KFbVpqHaUv3EsNaFXK9dpm_hiKs835Wp_mt57YZvlu0").getSheetByName("Processed");
  var occurrences = sheet.getRange(2, 8, 1000).getValues();
  //finding min
  var minCountries = {};
  var count = 0;
  var currMin = 1000000;
  for (var i = 0; i < occurrences.length; i++){
    if(parseInt(occurrences[i]) < currMin){
      currMin = occurrences[i];
    }
  }
  for (var j = 0; j < occurrences.length; j++){
    if(parseInt(occurrences[j]) == currMin){
      minCountries[count] = j;
      count++;
    }
  }
  for(var k = 0; k <count; k++){
    //sheet.getRange(k+1, 19).setValue(occurrences[minCountries[k]]);
    sheet.getRange(k+1, 19).setValue(sheet.getRange(minCountries[k], 7).getValue());
    
    sheet.getRange(k+1, 20).setValue(currMin);
  }
}

function makeSlide() {
  var name = SpreadsheetApp.openById("1KFbVpqHaUv3EsNaFXK9dpm_hiKs835Wp_mt57YZvlu0").getSheetByName("Processed").getRange(2, 11).getValue().toString();
  var sale = SpreadsheetApp.openById("1KFbVpqHaUv3EsNaFXK9dpm_hiKs835Wp_mt57YZvlu0").getSheetByName("Processed").getRange(2, 12).getValue().toString();
  
  var slides = SlidesApp.openById("1njNXGvwa-V9Db64tjipFIfDDgF1fFwxw-bhp0snAkYg").getSlides();
  var myslide = slides[0];
  var myshape = myslide.getShapes()[0];
  myshape.getText().setText("Today’s Top Seller is " + name + " with a total sold of $" + sale );
  
}
