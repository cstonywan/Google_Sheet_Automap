var ToppingOut = 42;      // The 封頂 column index
var namecolumnindex = 37  //The machine name column index
var mapmaxheight  = 31    //The left hand side map max height
var mapmaxwidth = 29      //The left hand side map max width
var s = SpreadsheetApp.getActive().getSheetByName('Basement Floor Plan');
var mapdata = s.getDataRange().getValues();
var infodata = s.getDataRange().getValues();
var lastrowindex = s.getLastRow();  //The sheet last row index
var Searchresult = 0;
var Typetable_startindex = 3; // Count table start row index 


var TypeClass =[];
/*加新機加呢到~~~*/ 
// var TypeClass = [
//   {type:"公司機",color:"#00ffff",num:0},   
//   {type:"4000",color:"#00ff00",num:0},
//   {type:"3000",color:"#ff9900",num:0}, 
//   {type:"壞機",color:"red",num:0},
//   {type:"分成",color:"#ffff00",num:0},
//   {type:"Monthly",color:"#0000ff",num:0},  
//   {type:"Season",color:"#ea01ff",num:0},
//   {type:"Hold",color:"#38761d",num:0},
// ]
 
var ShapeClass = [
  {shape: 4, optNumRows:2, optNumColumns:2 },
  {shape: '2v', optNumRows:2, optNumColumns:1 },
  {shape: '2h', optNumRows:1, optNumColumns:2 },
  {shape: '', optNumRows:1, optNumColumns:1 }
]

function Mainbasement(){
  TypeClass = [];
  Automapping();
  Creattypecaltable();
}

function Typeclasscreate(){
  var maxrownum = s.getRange("AW1:AW").getValues().filter(String).length;
  for(var i = 1; i <= maxrownum; i++){    
    TypeClass.push({type:s.getRange(i,Typeclasscolumnindex).getValue(), color:s.getRange(i,Typeclasscolumnindex).getBackground(),num:0})
    //console.log(s.getRange(i,Typeclasscolumnindex).getValue(),s.getRange(i,Typeclasscolumnindex).getBackground())
  }
}

function Automapping() { //main function
  Typeclasscreate();
  for (var i = 1; i <= mapmaxheight; i++) {
    for (var j = 1; j <= mapmaxwidth; j++) {      
        var x = i + 1;
        var y = j + 1;
        if (mapdata[i][j] && !isNaN(mapdata[i][j])){
          console.log("test index: "+"i:"+ i +" j:" + j + " " + mapdata[i][j]);
          console.log("map(int): "+"i:"+ i +" j:" + j + " " + mapdata[i][j] +"-> getToppingPrice: "+infodata[mapdata[i][j]][ToppingOut]);
          SetMapBackgroundcolour(infodata[mapdata[i][j]][ToppingOut], i,j)
          Changebacktoname(i,j)                  
          MergeBlockCell(x,y, infodata[mapdata[i][j]][46]);                      
        }            
        else{
          if (mapdata[i][j] && isNaN(mapdata[i][j] && mapdata[i][j] != 'Exit')){    
              var nameindex = GetNameindex(mapdata, mapdata[i][j]);
              SetMapBackgroundcolour(infodata[nameindex][ToppingOut], i,j)                
              MergeBlockCell(x,y, infodata[nameindex][46]);                        
          }
        }        
    }
  }    
}

function display(){
  var x =0 ;
  var y = 0;

    console.log(s.getRange(x,y).getValue(),s.getRange(x,y).getBackground())

}

// function SearchPrompt() {
//   var ui = SpreadsheetApp.getUi();
//   var result = ui.prompt("機台號碼/電話/姓名:",ui.ButtonSet.OK_CANCEL);
//   // Logger.log(result.getResponseText());
//   //  findinfodatarow(result.getResponseText());
//   return result.getResponseText();
// }

// function SearchEngine(){
//   Clearmapcolor();
//   //var value = s.getRange(1,namecolumnindex).getValue();
//   var value = SearchPrompt();
//   s.getRange(4,namecolumnindex).setValue(value);
//   s.getRange(4,namecolumnindex).setFontSize(14);
//   SetHorizontalAlignment(4,namecolumnindex);
  
//   if(!isNaN(value) && value.toString().length < 7){ //index        
//     SearchEnginebyindex(parseInt(value));     
//     Findinfodatarow(parseInt(value)+ 1);  
//   }
//   else{    
//         if(!isNaN(value) && value.toString().length == 8){ //phone
//             for (var i = 1; i < lastrowindex; i++) {  
//                 console.log("(mapdata[i][40]",i,mapdata[i][40])      
//                 if(mapdata[i][40].toString().length > 8){
//                     if(mapdata[i][40].toString().indexOf(value.toString()) > -1){                            
//                         SearchEnginebyindex(i);     
//                         Findinfodatarow(i+ 1);    
//                     }
//                 }
//                 else{
//                     if(mapdata[i][40] == value.toString()){                            
//                       SearchEnginebyindex(i);     
//                       Findinfodatarow(i + 1);    
//                     }
//                 }           
//             }
//         }
//         else{ //name        
//             if(isNaN(value)){
//                 for (var i = 1; i < lastrowindex; i++) {  
//                     if(value.toString().match(/[\u3400-\u9FBF]/)){
//                         //console.log("value",mapdata[i][39].toString())
//                         if(mapdata[i][39].toString().indexOf(value.toString()) > -1){                              
//                             SearchEnginebyindex(i);     
//                             Findinfodatarow(i + 1);    
//                         }
//                     }
//                     else{
//                       console.log("value",mapdata[i][39].toString() )                                                                                                            
//                         if(mapdata[i][39].toString().toLocaleLowerCase().indexOf(value.toString().toLocaleLowerCase()) > -1){                           
//                             SearchEnginebyindex(i);     
//                             Findinfodatarow(i+1);    
//                         }
//                     }            
//                 }
//             }
//         }
//     }
// }

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
 
// function SearchEnginebyindex(value){
//   for (var i = 1; i <= mapmaxheight; i++) {
//       for (var j = 1; j <= mapmaxwidth; j++) {          
//         if (mapdata[i][j] && !isNaN(mapdata[i][j])){
//           if(mapdata[i][j] == value){
//             console.log("in num",i,j)            
//             Setbackgroundcolor(i+1,j+1,"red")
//             s.getRange(5,36).setValue(i+1);
//             s.getRange(5,37).setValue(j+1);
//           }
//         }
//         else{
//           if (mapdata[i][j] && isNaN(mapdata[i][j])){    
//             if(mapdata[i][j] == value){
//               console.log("word",i,j)              
//               Setbackgroundcolor(i+1,j+1,"red")
//               s.getRange(5,36).setValue(i+1);
//               s.getRange(5,37).setValue(j+1);
//             }    
//           }
//         }        
//       }
//     }
// }
 
function Changebacktoname(i,j){
  if(mapdata[mapdata[i][j]][namecolumnindex] != ''){
    //console.log("change name: "+"i:"+ i +" j:" + j + " " + infodata[mapdata[i][j]][namecolumnindex]);
    s.getRange(i+1,j+1).activate().setValue(infodata[mapdata[i][j]][namecolumnindex]);
  }
}
 
// function GetNameindex(dbdata, mapdata){  
//   for (var i = 1; i <= lastrowindex; i++) {          
//     if(dbdata[i][namecolumnindex+1] == mapdata){      
//       return i;
//     }    
//   }
// }
 
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
 
// function Clearmapformat(){
//   for (var i = 2; i <= 14; i++) {
//     for (var j = 2; j <= 33; j++) {      
//       s.getRange(i,j).clearFormat();
//     }
//   }
// }
 
function Creattypecaltable(){
  var Tablestartrowindex = Typetable_startindex
  var Tablestartcolindex = 36
  var Totalsumoftype = 0

  Typeclasscreate();

  for(var a = 0; a < TypeClass.length;a++){      
    if(a == 0){
      TypeClass[0].num = Calnum('');
    }
    else{
      TypeClass[a].num = Calnum(TypeClass[a].type);
    } 
  }

  for(var a = 0; a < TypeClass.length;a++){     
    Totalsumoftype += TypeClass[a].num;
    console.log(TypeClass[a].type,TypeClass[a].color,TypeClass[a].num);
     for(var r = Tablestartrowindex; r < TypeClass.length; r ++){
      s.getRange(Tablestartrowindex+a,Tablestartcolindex).activate().setValue(TypeClass[a].type);
      s.getRange(Tablestartrowindex+a,Tablestartcolindex).setBackground(TypeClass[a].color);
      Settablestyle(Tablestartrowindex+a,Tablestartcolindex,14,"bold");   
      s.getRange(Tablestartrowindex+a,Tablestartcolindex+1).activate().setValue(TypeClass[a].num);  
      Settablestyle(Tablestartrowindex+a,Tablestartcolindex+1,12,"bold");
    }
  }
  s.getRange(Tablestartrowindex+TypeClass.length,Tablestartcolindex).activate().setValue('Total');
  Settablestyle(Tablestartrowindex+TypeClass.length,Tablestartcolindex,14,"bold");
  s.getRange(Tablestartrowindex+TypeClass.length,Tablestartcolindex+1).activate().setValue(Totalsumoftype);
  Settablestyle(Tablestartrowindex+TypeClass.length,Tablestartcolindex+1,12,"bold");
}
 
function Settablestyle(x,y,fontsize,fontstyle){
  s.getRange(x,y).setFontSize(fontsize);
  s.getRange(x,y).setFontWeight(fontstyle);
  SetHorizontalAlignment(x,y);
  Setboarder(x,y);
}

function Cleartypetable(){
  for(var i = Typetable_startindex; i < lastrowindex; i++){
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
 

// function GetRandomColor() {
//     var letters = 'BCDEF'.split('');
//     var color = '#';
//     for (var i = 0; i < 6; i++ ) {
//         color += letters[Math.floor(Math.random() * letters.length)];
//     }
//     return color;    
// }

function Colorcodecheck(){
  var x = 39;
  var y = 1;
  s.getRange(x,y).setValue(s.getRange(x,y).getBackground());
}