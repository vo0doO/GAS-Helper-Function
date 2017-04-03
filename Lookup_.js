//~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`
//--//Dependent on isEmpty_()
// Script Look-up
/*
Преимущество этого скрипта: 
-Такие листы google не будут постоянно выполнять поиск данных, которые не изменяются при использовании этой функции, поскольку она задана с жесткими значениями, пока сценарий не будет запущен снова. 
-В отличие от Vlookup вы можете просмотреть его для справочных данных в любом столбце в строке. Не должно быть в первом столбце, чтобы он работал как Vlookup. 
-Вы можете вернуть Lookup to Memory для дальнейшей обработки с помощью других функций

Search terms:
 - Google App Script / GAS / Javascript 
 - Vlookup / lookup / Search / query

Usage:
var SheetinfoArray    = SpreadsheetApp.openById(SheetID).getSheetByName('Sheet1').getRange('J2:J').getValues();

Lookup_(SheetinfoArray,"Sheet1!A:B",0,[1],"Sheet1!I1","n","y","n");
//or
Lookup_(Sheetinfo,"Sheet1!A:B",0,[1],"return","n","n","y");
//or
Lookup_(SheetinfoArray,"Sheet1!A:B",0,[0,1],"return","n","n","y");
//or
Lookup_(Sheetinfo,"Sheet1!A:B",1,[1,3,0,2],"return","y","n","n");
//or
Lookup_("female","Sheet1!A:G",4,[2],"Database!A1","y","y","y");
//or
Lookup_(Sheetinfo,LocationsArr,4,[0],"return","y","y","y");

-Загрузить все номера мест из J2: J в переменную 30 
- ищет номера мест в столбце 0 справочного листа и диапазон, например: «Данные! A: G» 31 
--- возвращает результаты в столбец 3 целевого листа и диапазон, например, «test! A1» или «1,1»,


Parameters Explaination:
"Search_Key" - Может быть одной ячейкой или массивом для одновременного поиска нескольких объектов

"RefSheetRange" - Справочный источник информации. Может быть локальной ссылкой на лист и диапазоном или массивом данных из переменной.

"SearchKey_RefMatch_IndexOffSet" - В каком столбце информации вы ссылаетесь на данные «Search_Key» на «RefSheetRange»

"IndexOffSetForReturn" - После того, как найдено совпадение 'Search_Key', какие столбцы данных будут возвращены из 'RefSheetRange'

"SetSheetRange" - Куда вы собираетесь поместить выбранную информацию из 'RefSheetRange', совпадающую с 'Search_Key' ИЛИ ​​вы можете использовать 'return', и когда функция завершится, она вернется, чтобы вы могли вывести функцию в переменную

"ReturnMultiResults" - Если 'Y' Говорят, что 'Search_Key' является 'NW', и вы хотите найти каждое хранилище в цепочке, которая попадает под северо-запад в ваш набор данных. Таким образом, объявление о том, что 'Y' не остановится после того, как оно найдет первое совпадение, продолжит поиск остальных данных

"Add_Note" - Если «Y» вы устанавливаете результаты в электронную таблицу и не возвращаете ее в память, тогда она будет устанавливать первую ячейку в «SetSheetRange» с примечанием о том, что и когда.

"Has_NAs" - Если 'Y', он помещает '# N / A' в столбец, где он не нашел данные для 'Search_Key', иначе он оставит столбец пустым.

*/

function Lookup_(Search_Key,RefSheetRange,SearchKey_RefMatch_IndexOffSet,IndexOffSetForReturn,SetSheetRange,ReturnMultiResults,Add_Note,Has_NAs)   
{
  if(/^y$/i.test(Has_NAs))
  {
    var NALoad = "#N/A";
  }
  else
  {
    var NALoad = "";
  }
  
  if(Object.prototype.toString.call(Search_Key) === '[object String]')
  {
    var Search_Key = new Array(Search_Key);
  }
  
  if(Object.prototype.toString.call(IndexOffSetForReturn) === '[object Number]')
  {
    var IndexOffSetForReturn = new Array(IndexOffSetForReturn.toString());
  }
  
  if(Object.prototype.toString.call(RefSheetRange) === '[object String]')
  {
    var RefSheetRangeArr = RefSheetRange.split("!");
    var Ref_Sheet = RefSheetRangeArr[0];
    var Ref_Range = RefSheetRangeArr[1];
    var data = SpreadsheetApp.getActive().getSheetByName(Ref_Sheet).getRange(Ref_Range).getValues();         //Syncs sheet by name and range into var
  }
  
  if(Object.prototype.toString.call(RefSheetRange) === '[object Array]')
  {
    var data = RefSheetRange;
  }
  
  if(!/^return$/i.test(SetSheetRange))
  {
  var SetSheetRangeArr = SetSheetRange.split("!");
  var Set_Sheet = SetSheetRangeArr[0];
  var Set_Range = SetSheetRangeArr[1];
  var RowVal = SpreadsheetApp.getActive().getSheetByName(Set_Sheet).getRange(Set_Range).getRow();
  var ColVal = SpreadsheetApp.getActive().getSheetByName(Set_Sheet).getRange(Set_Range).getColumn();
  }
  
  var twoDimensionalArray = [];
  for (var i = 0, Il=Search_Key.length; i<Il; i++)                                                         // i = number of rows to index and search  
  {
    var Sending = [];                                                                                      //Making a Blank Array
    var newArray = [];                                                                                     //Making a Blank Array
    var Found ="";
    for (var nn=0, NNL=data.length; nn<NNL; nn++)                                                                 //nn = will be the number of row that the data is found at
    {
      if(Found==1 && /^n$/i.test(ReturnMultiResults))                                                                                         //if statement for found if found = 1 it will to stop all other logic in nn loop from running
      {
        break;                                                                                             //Breaking nn loop once found
      }
      if (data[nn][SearchKey_RefMatch_IndexOffSet]==Search_Key[i])                                              //if statement is triggered when the search_key is found.
      {
        var newArray = [];
        for (var cc=0, CCL=IndexOffSetForReturn.length; cc<CCL; cc++)                                         //cc = numbers of columns to referance
        {
          var iosr = IndexOffSetForReturn[cc];                                                             //Loading the value of current cc
          var Sending = data[nn][iosr];                                                                    //Loading data of Level nn offset by value of cc
          if(isEmpty_(Sending))                                                                            //if statement for if one of the returned Column level cells are blank
          {
            var Sending =  NALoad;                                                                           //Sets #N/A on all column levels that are blank
          }
          if (CCL>1)                                                                                       //if statement for multi-Column returns
          {
            newArray.push(Sending);
            if(CCL-1 == cc)                                                                                //if statement for pulling all columns into larger array
            {
              twoDimensionalArray.push(newArray);
              var Found = 1;                                                                              //Modifying found to 1 if found to stop all other logic in nn loop
              break;                                                                                      //Breaking cc loop once found
            }
          }
          else if (CCL<=1)                                                                                 //if statement for single-Column returns
          {
            twoDimensionalArray.push(Sending);
            var Found = 1;                                                                                 //Modifying found to 1 if found to stop all other logic in nn loop
            break;                                                                                         //Breaking cc loop once found
          }
        }
      }
      if(NNL-1==nn && isEmpty_(Sending))                                                             //following if statement is for if the current item in lookup array is not found.  Nessessary for data structure.
      {
        for(var na=0,NAL=IndexOffSetForReturn.length;na<NAL;na++)                                          //looping for the number of columns to place "#N/A" in to preserve data structure
        {
          if (NAL<=1)                                                                                      //checks to see if it's a single column return
          {
            var Sending = NALoad;
            twoDimensionalArray.push(Sending);
          }
          else if (NAL>1)                                                                                  //checks to see if it's a Multi column return
          {
            var Sending = NALoad;
            newArray.push(Sending);
          }
        }
        if (NAL>1)                                                                                         //checks to see if it's a Multi column return
        {
          twoDimensionalArray.push(newArray);  
        }
      }
    }
  }
  if(!/^return$/i.test(SetSheetRange))
  {
    if (CCL<=1)                                                                                            //checks to see if it's a single column return for running setValue
    {
      var singleArrayForm = [];
      for (var l = 0,lL=twoDimensionalArray.length; l<lL; l++)                                                          //Builds 2d Looping-Array to allow choosing of columns at a future point
      {
        singleArrayForm.push([twoDimensionalArray[l]]);
      }
      SpreadsheetApp.getActive().getSheetByName(Set_Sheet).getRange(RowVal,ColVal,singleArrayForm.length,singleArrayForm[0].length).setValues(singleArrayForm);
    }
    if (CCL>1)                                                                                             //checks to see if it's a multi column return for running setValues
    {
      SpreadsheetApp.getActive().getSheetByName(Set_Sheet).getRange(RowVal,ColVal,twoDimensionalArray.length,twoDimensionalArray[0].length).setValues(twoDimensionalArray);
    }
    if(/^y$/i.test(Add_Note))
    {
      if(Object.prototype.toString.call(RefSheetRange) === '[object Array]')
      {
        SpreadsheetApp.getActive().getSheetByName(Set_Sheet).getRange(RowVal,ColVal,1,1).setNote("VLookup Script Ran On: " + Utilities.formatDate(new Date(), "PST", "MM-dd-yyyy hh:mm a") + "\nRange: Origin Variable" );      
      }
      if(Object.prototype.toString.call(RefSheetRange) === '[object String]')
      {
        SpreadsheetApp.getActive().getSheetByName(Set_Sheet).getRange(RowVal,ColVal,1,1).setNote("VLookup Script Ran On: " + Utilities.formatDate(new Date(), "PST", "MM-dd-yyyy hh:mm a") + "\nRange: " + RefSheetRange);      
      }
    }
    SpreadsheetApp.flush();
    SpreadsheetApp.getActiveSpreadsheet().toast('At: '+SetSheetRange,'Lookup Completed:');
  }
  if(/^return$/i.test(SetSheetRange))
  {
    return twoDimensionalArray
  }
}
//~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`~,~`
