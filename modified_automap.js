/*
 By Tony Wan on 2022-04-30
*/
var ToppingOut = 42;      // The 封頂 column index
var namecolumnindex = 37  //The machine name column index
var mapmaxheight  = 13    //The left hand side map max height
var mapmaxwidth = 33      //The left hand side map max width
var s = SpreadsheetApp.getActive().getSheetByName('CGA GF FLOOR PLAN');
var mapdata = s.getDataRange().getValues();
var infodata = s.getDataRange().getValues();
var lastrowindex = s.getLastRow();  //The sheet last row index 
var Searchresult = 0;
 
  /*var TypeClass = [
    ["公司機","#00ffff"],
    ["Lease","#02ff01"],
    ["Hold","#bb00ff"],
    ["broke","#ff0000"],
    ["Exchange","#ff4d00"],
    ["share","#ffff02"],
    ["Smallb","#1f5ded"]  
  ] */
 
var TypeClass = [
  {type:"公司機",color:"#00ffff"},
  {type:"4000租",color:"#02ff01"},
  {type:"Hold",color:"#bb00ff"},
  {type:"壞機",color:"pink"},
  {type:"Exchange",color:"#76b5c5"},
  {type:"分成機",color:"#ffff02"},
  {type:"細B",color:"#1f5ded"},
]
 
var ShapeClass = [
  {shape: 4, optNumRows:2, optNumColumns:2 },
  {shape: '2v', optNumRows:2, optNumColumns:1 },
  {shape: '2h', optNumRows:1, optNumColumns:2 },
  {shape: '', optNumRows:1, optNumColumns:1 }
]
 
function Automapping() { //main function
 
  for (var i = 1; i <= mapmaxheight; i++) {
    for (var j = 1; j <= mapmaxwidth; j++) {      
        var x = i + 1;
        var y = j + 1;
        if (mapdata[i][j] && !isNaN(mapdata[i][j])){
          //console.log("test index: "+"i:"+ i +" j:" + j + " " + mapdata[i][j]);
          //console.log("map(int): "+"i:"+ i +" j:" + j + " " + mapdata[i][j] +"-> getToppingPrice: "+infodata[mapdata[i][j]][ToppingOut]);
          SetMapBackgroundcolour(infodata[mapdata[i][j]][ToppingOut], i,j)
          Changebacktoname(i,j)                  
          MergeBlockCell(x,y, infodata[mapdata[i][j]][46]);                      
        }            
        else{
          if (mapdata[i][j] && isNaN(mapdata[i][j])){    
              var nameindex = GetNameindex(mapdata, mapdata[i][j]);
              SetMapBackgroundcolour(infodata[nameindex][ToppingOut], i,j)                
              MergeBlockCell(x,y, infodata[nameindex][46]);                        
          }
        }        
    }
  }    
}

function SearchPrompt() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt("機台號碼/電話/姓名:",ui.ButtonSet.OK_CANCEL);
  // Logger.log(result.getResponseText());
  //  findinfodatarow(result.getResponseText());
  return result.getResponseText();
}

function SearchEngine(){
  Clearmapcolor();
  //var value = s.getRange(1,namecolumnindex).getValue();
  var value = SearchPrompt();
  s.getRange(4,namecolumnindex).setValue(value);
  s.getRange(4,namecolumnindex).setFontSize(14);
  SetHorizontalAlignment(4,namecolumnindex);
  
  if(!isNaN(value) && value.toString().length < 7){ //check for index        
    SearchEnginebyindex(parseInt(value));     
    Findinfodatarow(parseInt(value)+ 1);  
  }
  else{    
        if(!isNaN(value) && value.toString().length == 8){ //phone
            for (var i = 1; i < lastrowindex; i++) {  
                //console.log("(mapdata[i][40]",i,mapdata[i][40])      
                if(mapdata[i][40].toString().length > 8){
                    if(mapdata[i][40].toString().indexOf(value.toString()) > -1){                            
                        SearchEnginebyindex(i);     
                        Findinfodatarow(i+ 1);    
                    }
                }
                else{
                    if(mapdata[i][40] == value.toString()){                            
                      SearchEnginebyindex(i);     
                      Findinfodatarow(i + 1);    
                    }
                }           
            }
        }
        else{ //name        
            if(isNaN(value)){
                for (var i = 1; i < lastrowindex; i++) {  
                    if(value.toString().match(/[\u3400-\u9FBF]/)){
                        //console.log("value",mapdata[i][39].toString())
                        if(mapdata[i][39].toString().indexOf(value.toString()) > -1){                              
                            SearchEnginebyindex(i);     
                            Findinfodatarow(i + 1);    
                        }
                    }
                    else{
                      //console.log("value",mapdata[i][39].toString() )                                                                                                            
                        if(mapdata[i][39].toString().toLocaleLowerCase().indexOf(value.toString().toLocaleLowerCase()) > -1){                           
                            SearchEnginebyindex(i);     
                            Findinfodatarow(i+1);    
                        }
                    }            
                }
            }
        }
    }
}

function Cleartablebg(){  
  for (var i = 1; i < lastrowindex; i++) {  
    for(var j = 39; j<48; j++){
      s.getRange(i,j).setBackground("white");
    }
  }
  //Automapping();
}
function Findinfodatarow(x){  
    Cleartablebg();
    for(var j = 39; j<48; j++){
        Setbackgroundcolor(x,j,"#c9fbff");
        //console.log("x: ",x);
    }
}
 
function SearchEnginebyindex(value){
  for (var i = 1; i <= mapmaxheight; i++) {
      for (var j = 1; j <= mapmaxwidth; j++) {          
        if (mapdata[i][j] && !isNaN(mapdata[i][j])){
          if(mapdata[i][j] == value){
            //console.log("in num",i,j)            
            Setbackgroundcolor(i+1,j+1,"red")
            s.getRange(5,36).setValue(i+1);
            s.getRange(5,37).setValue(j+1);
          }
        }
        else{
          if (mapdata[i][j] && isNaN(mapdata[i][j])){    
            if(mapdata[i][j] == value){
              //console.log("word",i,j)              
              Setbackgroundcolor(i+1,j+1,"red")
              s.getRange(5,36).setValue(i+1);
              s.getRange(5,37).setValue(j+1);
            }    
          }
        }        
      }
    }
}
 
function Changebacktoname(i,j){
  if(mapdata[mapdata[i][j]][namecolumnindex] != ''){
    //console.log("change name: "+"i:"+ i +" j:" + j + " " + infodata[mapdata[i][j]][namecolumnindex]);
    s.getRange(i+1,j+1).activate().setValue(infodata[mapdata[i][j]][namecolumnindex]);
  }
}
 
function GetNameindex(dbdata, mapdata){  
  for (var i = 1; i <= lastrowindex; i++) {          
    if(dbdata[i][namecolumnindex+1] == mapdata){      
      return i;
    }    
  }
}
 
function SetMapBackgroundcolour(mapvalue, i, j){
  var x = i + 1
  var y = j + 1  
 
  if(!mapvalue){
    Setbackgroundcolor(x,y,TypeClass[0].color);
    TypeClass[0].num++;
  }
  else{
    for(var i = 0; i <= TypeClass.length-1; i++){
      if(mapvalue == TypeClass[i].type){
          Setbackgroundcolor(x,y,TypeClass[i].color);
          TypeClass[i].num++;
      }
    }        
  }
}
 
function MergeBlockCell(x,y,type){  
  for(var i = 0; i<ShapeClass.length; i++){
      if(ShapeClass[i].shape == type){
        s.getRange(x,y,ShapeClass[i].optNumRows,ShapeClass[i].optNumColumns).merge(); //getRange(row, column, optNumRows, optNumColumns)  
      }
  }  
  SetHorizontalAlignment(x,y);  
  s.getRange(x,y).setBorder(true, true, true, true, true, true, "Black", SpreadsheetApp.BorderStyle.SOLID_THICK);
  s.getRange(x,y).setFontWeight("bold");
}
 
function SetHorizontalAlignment(x,y) {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var s= ss.getActiveSheet()
  var lr = lastrowindex
  var r= s.getRange(x, y, lr,2)
  var set=r.setHorizontalAlignment("center")
}
 
function Clearmapcolor(){
  for (var i = 1; i <= mapmaxheight; i++) {
    for (var j = 1; j <= mapmaxwidth; j++) {      
      s.getRange(i,j).setBackground("white");
    }
  }
}
 
function Clearmapformat(){
  for (var i = 2; i <= 14; i++) {
    for (var j = 2; j <= 33; j++) {      
      s.getRange(i,j).clearFormat();
    }
  }
}
 
function Creattypecaltable(){
  var Typearray = ['公司機']
  var Tablearr = []
  var Tablestartrowindex = 7
  var Tablestartcolindex = 36
  var Totalsumoftype = 0
 
  for(var i = 1;  i < lastrowindex; i++){      
    if(!Typearray.includes(infodata[i][ToppingOut]) && infodata[i][ToppingOut] != ''){
      Typearray.push(infodata[i][ToppingOut])     
    }    
  }  
 
  Tablearr.push(['','Total'])
  Tablearr.push([Typearray[0],Calnum('')])
  for(var a = 1;  a < Typearray.length; a++){                  
    Tablearr.push([Typearray[a],Calnum(Typearray[a])]);     
  }
 
  for(var i = 1; i < Tablearr.length;i++){
    Totalsumoftype += Tablearr[i][1];    
  }
  Tablearr.push(["Total",Totalsumoftype])    
  for(var i = 0; i < Tablearr.length; i++){
   
    s.getRange(Tablestartrowindex+i,Tablestartcolindex).activate().setValue(Tablearr[i][0]);  
    if(i<TypeClass.length){
      //console.log(Tablestartrowindex+i+1,TypeClass[i].color)
      Setbackgroundcolor(Tablestartrowindex+i+1,Tablestartcolindex,TypeClass[i].color)
    }
    s.getRange(Tablestartrowindex+i,Tablestartcolindex+1).activate().setValue(Tablearr[i][1]);
    Settablestyle(Tablestartrowindex+i,Tablestartcolindex,14,"bold");
    Settablestyle(Tablestartrowindex+i,Tablestartcolindex+1,12,"bold");
  }  
}
 
function Settablestyle(x,y,fontsize,fontstyle){
  s.getRange(x,y).setFontSize(fontsize);
  s.getRange(x,y).setFontWeight(fontstyle);
  SetHorizontalAlignment(x,y);
  Setboarder(x,y);
}
function Cleartypetable(){
  for(var i = 7; i < lastrowindex; i++){
     Cleartable(i,36);
     Cleartable(i,37);
  }
}
function Cleartable(i,j){    
  s.getRange(i,j).clearContent();
  s.getRange(i,j).clearFormat();
}
 
function Calnum(type){
  var count = 0;
  for(var i = 0; i < lastrowindex; i++){      
    if(infodata[i][ToppingOut] == type){
      count++;
    }
  }
  return count
}
 
function Setboarder(x,y){
  s.getRange(x,y).setBorder(true, true, true, true, false, false, "Black", SpreadsheetApp.BorderStyle.SOLID);
}
 
function Setbackgroundcolor(x,y,color){
  s.getRange(x,y).setBackground(color);
}
 

function GetRandomColor() {
    var letters = 'BCDEF'.split('');
    var color = '#';
    for (var i = 0; i < 6; i++ ) {
        color += letters[Math.floor(Math.random() * letters.length)];
    }
    return color;
  }