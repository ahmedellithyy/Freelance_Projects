function dataExtraction(ticker,date)
{
  ticker = encodeURI(ticker);
  var requestOptions = {
    'method': 'GET',
    'contentType':'application/json',
    'headers':{
    'x-api-key': "API_KEY" // you get it from your yahoo mail
    }
  };
  var response = UrlFetchApp.fetch("https://rest.yahoofinanceapi.com/v7/finance/options/"+ticker+"?date="+date,requestOptions);  
  var chain = JSON.parse(response.getContentText());
  return chain.optionChain.result[0].options[0];

}
//OPTIONSFUNCTION(“Enter Ticker”, “Enter Strike Price”, “Enter Expiry Date”,"Calls or Puts")
/** 
@customfunction
*/
function optionRow(ticker,price,date,cp){
  var unixDate = parseInt((new Date(date+" 02:00:00") / 1000).toFixed(0))
  var extractedData = dataExtraction(ticker,unixDate)
  var arrangedData = dataManipulation(extractedData)
  var calls = arrangedData["calls"]
  var puts = arrangedData["puts"]
  if(typeof ticker != "string")  throw "Expected a string of the ticker but got an input of another type.";
  if(typeof price != "number")  throw "Expected a number of the strike price but got an input of another type.";

  if(cp.toUpperCase() == "P"){
    for(let i = 0; i < puts.length;i++){
      if(price == puts[i][2]){
        var row = []
        row.push(puts[i])
        return row
      }
    }
    throw "No Matching Value for this date"
  }
  else if (cp.toUpperCase() == "C"){
    for(let i = 0; i < calls.length;i++){
      if(price == calls[i][2]){
        var row = []
        row.push(calls[i])
        return row
      }
    }
    throw "No Matching Value for this data"

  }
  else{
    throw "Expected just two values P for puts, C for calls"
  }
  
}


function dataManipulation(data){
  var callsRows = [];
  var putsRows = []
  for(var i =0; i< data["calls"].length;i++){
    
    callsRows.push([data["calls"][i].contractSymbol,""/*lastTradeDate*/,data["calls"][i].strike,data["calls"][i].lastPrice,data["calls"][i].bid,data["calls"][i].ask,data["calls"][i].change,data ["calls"][i].percentChange,data["calls"][i].volume,data["calls"][i].openInterest,data["calls"][i].impliedVolatility
    ])

    var date = new Date(data["calls"][i].lastTradeDate * 1000).toISOString().slice(0, 19).replace('T', ' ')
    callsRows[i][1] =new Date(date).toLocaleString('en-US', { timeZone: 'America/New_York' });

  }
  for(var i =0; i< data["puts"].length;i++){
    
    putsRows.push([data["puts"][i].contractSymbol,""/*lastTradeDate*/,data["puts"][i].strike,data["puts"][i].lastPrice,data["puts"][i].bid,data["puts"][i].ask,data["puts"][i].change,data["puts"][i].percentChange,data["puts"][i].volume,data["puts"][i].openInterest,data["puts"][i].impliedVolatility
    ])
    var date = new Date(data["puts"][i].lastTradeDate * 1000).toISOString().slice(0, 19).replace('T', ' ')
    putsRows[i][1] =new Date(date).toLocaleString('en-US', { timeZone: 'America/New_York' });

  }
  
   var allData = {
     "calls": callsRows,
     "puts":putsRows
   }
   return allData
}
function dataUpdate(data){
  var callsRows = [];
  var putsRows = []
  for(var i =0; i< data["calls"].length;i++){
    
    callsRows.push([data["calls"][i].contractSymbol,""/*lastTradeDate*/,data["calls"][i].strike,data["calls"][i].lastPrice,data["calls"][i].bid,data["calls"][i].ask,data["calls"][i].change,data ["calls"][i].percentChange,data["calls"][i].volume,data["calls"][i].openInterest,data["calls"][i].impliedVolatility
    ])

    var date = new Date(data["calls"][i].lastTradeDate * 1000).toISOString().slice(0, 19).replace('T', ' ')
    callsRows[i][1] = date
  }
  var callsLiveSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Calls Live Data") // write sheet name 
  update = callsLiveSheet.getRange(3,1,callsRows.length,11).setValues(callsRows);


  for(var i =0; i< data["puts"].length;i++){
    
    putsRows.push([data["puts"][i].contractSymbol,""/*lastTradeDate*/,data["puts"][i].strike,data["puts"][i].lastPrice,data["puts"][i].bid,data["puts"][i].ask,data["puts"][i].change,data["puts"][i].percentChange,data["puts"][i].volume,data["puts"][i].openInterest,data["puts"][i].impliedVolatility
    ])
    var date = new Date(data["puts"][i].lastTradeDate * 1000).toISOString().slice(0, 19).replace('T', ' ')
    putsRows[i][1] = date

  }
    var putsLiveSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Puts Live Data") // write sheet name 
    update = putsLiveSheet.getRange(3,1,putsRows.length,11).setValues(putsRows);
  
}

function manualMain(ticker)
{
  var data = dataExtraction(ticker)
  dataUpdate(data)
    var interface = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Interface")
    var now = new Date()
    updateLiveSymbol = interface.getRange(3,3).setValue(ticker);
    updateLiveLastUpdate = interface.getRange(3,4).setValue(now);
}

/*function liveMain(){
  var interface = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Interface")
  var symbol = interface.getRange(3,1).getValue();
  var data = dataExtraction(symbol)
  dataUpdate(data)
  var now = new Date()
  updateLiveLastUpdate = interface.getRange(3,2).setValue(now)
}*/

