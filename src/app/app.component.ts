import { Component } from '@angular/core';

@Component({
  selector: 'crosstab',
  template: `
              <gc-spread-sheets (workbookInitialized)="workbookInit($event)">
              <div id="crosstabHeader"></div>
              </gc-spread-sheets>`,
})
export class AppComponent {
  workbookInit(args: any) {

    let spread = args.spread;
    let sheet = spread.getActiveSheet();
    //spread.options.tabNavigationVisible = false;
    spread.options.newTabVisible = false;//隐藏新增tab选项
    //spread.options.showHorizontalScrollbar=false;//隐藏水平滚动条
    //spread.options.showVerticalScrollbar=false;//隐藏垂直滚动条
    spread.options.tabStripVisible =false;//隐藏下方tab导航
    spread.options.grayAreaBackColor='white';//设置表单灰色区域背景色
		//spread.invalidateLayout();
		//spread.repaint();
    //sheet.options.isProtected = true;//用户不可编辑
    GC.Spread.Common.CultureManager.culture("zh-cn");
    let datas = getSimpleDataSet(5);
    var jsonTarget: any[] = [];
    //格式方案
    var allFormatter = [{
      "id": "01",
      "name": "01",
      "row": ["row1","row2"],
      "col": ["col1","col2"],
      "value": "value1,value2",
      "mainTitle":{"name":"交叉表主标题","level":0,"style":"text-align: center;font-size: 10px;font-family: 宋体;font-weight: bold;font-style: italic;height: 10px;vertical-align: middle;display: table-cell;"},
      "subTitles":[
                    {"name":"交叉表副标题1","level":1,"style":"text-align: center;font-size: 10px;font-family: 宋体;font-weight: bold;font-style: italic;height: 10px;vertical-align: middle;display: table-cell;"},
                    {"name":"交叉表副标题2","level":3,"style":"text-align: center;font-size: 10px;font-family: 宋体;font-weight: bold;font-style: italic;height: 10px;vertical-align: middle;display: table-cell;"},
                    {"name":"交叉表副标题3.1","level":2,"style":"text-align: left;font-size: 10px;font-family: 宋体;font-weight: bold;font-style: italic;height: 10px;vertical-align: middle;display: table-cell;"},
                    {"name":"交叉表副标题3.2","level":4,"style":"text-align: right;font-size: 10px;font-family: 宋体;font-weight: bold;font-style: italic;height: 10px;vertical-align: middle;display: table-cell;"}],
       "gridSetting":{"rowSum":true,"colSum":true}
    }];
    //图标设置
    let gridSetting = allFormatter.find(x => x.id == "01").gridSetting;
    //基于后端返回对象获取行维度所有的数据
    let rowInfos = allFormatter.find(x => x.id == "01").row;
    let rowInfoWhere='',rowSortInfoWhere='',rowSumInfoWhere='';
    for(let rowInfo of rowInfos)
    {
      rowInfoWhere=rowInfoWhere+'item.'+rowInfo+' == currentValue.'+rowInfo+' && ';
      rowSumInfoWhere=rowSumInfoWhere+'item.'+rowInfo+' == sum.'+rowInfo+' && ';
      rowSortInfoWhere=rowSortInfoWhere+'x.'+rowInfo+'.localeCompare(y.'+rowInfo+') || ';
    }
    rowInfoWhere=rowInfoWhere.substr(0,rowInfoWhere.length-3);
    rowSumInfoWhere=rowSumInfoWhere.substr(0,rowSumInfoWhere.length-3);
    rowSortInfoWhere=rowSortInfoWhere.substr(0,rowSortInfoWhere.length-3);

    var rowDismension = datas.reduce((previousValue, currentValue) => {
    let  b={}
    if (!previousValue.find(item => eval('('+rowInfoWhere+')'))) {
      for(let rowInfo of rowInfos)
      {
        b[rowInfo]=currentValue[rowInfo];
      }
      previousValue.push(b);
    };
    return previousValue;
  }, []);
    rowDismension.sort((x,y)=>eval('('+rowSortInfoWhere+')'));
    
    /*
    if(1==1){
          for(let a=1;a<rowInfos.length;a++){
            var sum={};
            for(let b=a;b>0;b--){
              sum[rowInfos[0]]=currentValue[rowInfos[0]];
            }
            sum[rowInfos[a]]='小计';
          }
          if (!previousValue.find(item => eval('('+rowSumInfoWhere+')'))) {
            previousValue.push(sum);
          }
        }
        */
    //基于后端返回对象获取列维度所有的数据
    let colInfos = allFormatter.find(x => x.id == "01").col;
    let colInfoWhere='',colSortInfoWhere='',colSumInfoWhere='';
    for(let colInfo of colInfos)
    {
      colInfoWhere=colInfoWhere+'item.'+colInfo+' == currentValue.'+colInfo+' && ';
      colSortInfoWhere=colSortInfoWhere+'x.'+colInfo+'.localeCompare(y.'+colInfo+') || ';
    }
    colInfoWhere=colInfoWhere.substr(0,colInfoWhere.length-3);
    colSortInfoWhere=colSortInfoWhere.substr(0,colSortInfoWhere.length-3);
    colSumInfoWhere=colInfoWhere.replace(/currentValue/g,'sum');
    let colDismension = datas.reduce((previousValue, currentValue) => {
    let  b={}
    if (!previousValue.find(item => eval('('+colInfoWhere+')'))) {
      for(let colInfo of colInfos)
      {
        b[colInfo]=currentValue[colInfo];
      }
      previousValue.push(b);
    };
    return previousValue;
    }, []);
    colDismension.sort((x,y)=>eval('('+colSortInfoWhere+')'));
    //最后一行添加小计
    var first:any=null;
    rowDismension.forEach((value ,index)=>{
      for(let a=1;a<rowInfos.length;a++){
            var temp={};
            for(let b=0;b<a;b++){
              temp[rowInfos[b]]=value[rowInfos[b]];
            }
            temp[rowInfos[a]]='小计';
            if(first==null){
              first=temp;
            }
            if(JSON.stringify(first)!=JSON.stringify(temp)){
              rowDismension.splice(index,0,first);
              first=temp;
            }
          }
    });
    if(first!=null){
      rowDismension.splice(rowDismension.length,0,first);
    }
    //最后一列添加小计
    first=null;
    colDismension.forEach((value ,index)=>{
      for(let a=1;a<colInfos.length;a++){
            var temp={};
            for(let b=0;b<a;b++){
              temp[colInfos[b]]=value[colInfos[b]];
            }
            temp[colInfos[b]]='小计';
            if(first==null){
              first=temp;
            }
            if(JSON.stringify(first)!=JSON.stringify(temp)){
              colDismension.splice(index,0,first);
              first=temp;
            }
          }
    });
    if(first!=null){
      colDismension.splice(colDismension.length,0,first);
    }
    //基于格式方案获取值列对象
    let valDismension = allFormatter.find(x => x.id == "01").value.split(',')
    let rowHeadersColumnCount = Object.keys(rowDismension[0]).length;//行维度数量(所占的列数)
    let columnHeaderRowCount = Object.keys(colDismension[0]).length + 1;//列维度数量（所占的行数)
    let valueHeaderRowCount = valDismension.length;//值维度数量
    let rowCount = getJsonLength(rowDismension);//行数
    let columnCount = getJsonLength(colDismension) * valueHeaderRowCount;//列数
    sheet.suspendPaint();
    sheet.setRowCount(rowCount);
    sheet.setColumnCount(columnCount);
    sheet.setRowCount(columnHeaderRowCount, GC.Spread.Sheets.SheetArea.colHeader);
    sheet.setColumnCount(rowHeadersColumnCount, GC.Spread.Sheets.SheetArea.rowHeader);
    

    //设置主标题
    let mainTitle = allFormatter.find(x => x.id == "01").mainTitle;
    if(mainTitle){
      let str ='<div id="mainTitle" style=\"'+mainTitle.style+'width:'+document.getElementById("crosstabHeader").clientWidth+'px\">'+mainTitle.name+'</div><br>';
      document.getElementById("crosstabHeader").innerHTML+=str;
    }
    //设置副标题
    let subTitles = allFormatter.find(x => x.id == "01").subTitles;
    if(subTitles){
      subTitles.sort((a,b)=>{return a.level-b.level}).forEach((value)=>{
        let str ='<div id="subTitle'+value.level+'" style=\"'+value.style+'width:'+document.getElementById("crosstabHeader").clientWidth+'px\">'+value.name+'</div><br>';
        document.getElementById("crosstabHeader").innerHTML+=str;
      })
    }
    /*
    let slicerDatas:any =[];

    rowDismension.forEach((data)=>{
      let slicerData=[];
      for(let value in data){
        slicerData.push(data[value]);
      }
      slicerDatas.push(slicerData);
    });
    var slicerData = new GC.Spread.Slicers.GeneralSlicerData(slicerDatas,rowInfos);
    var onFiltered = slicerData.onFiltered;
    slicerData.onFiltered = function () {
        onFiltered.call(slicerData);
        refreshList(slicerData);
    }
    var nameSlicer = new GC.Spread.Sheets.Slicers.ItemSlicer("row1", slicerData, "row1");
    nameSlicer.height(200);
    nameSlicer.width(180)
    nameSlicer.columnCount(2);
    document.getElementById('row1').appendChild(nameSlicer.getDOMElement());

    var classSlicer = new GC.Spread.Sheets.Slicers.ItemSlicer("row2", slicerData, "row2");
    classSlicer.height(200);
    classSlicer.width(180)
    document.getElementById('row2').appendChild(classSlicer.getDOMElement());
    */
    //设置行标题
    let sumIndex=1;//记录起始位置
    let sumAllRow=[];//特殊处理行号
    for (var r = 0; r < rowCount; r++) {
      for (var c = 0; c < rowHeadersColumnCount; c++) {
        let rowName = Object.keys(rowDismension[0])[c];
        sheet.setValue(r, c, rowDismension[r][rowName], GC.Spread.Sheets.SheetArea.rowHeader)
        //最后一行添加小计
        if(rowDismension[r][rowName]=='小计'){
          for(let colCal = 0; colCal <= columnCount; colCal++){
            sheet.setFormula(r, colCal, "SUBTOTAL(109,"+numToString(colCal)+  sumIndex + ":"+numToString(colCal) + r + ")");
            sumAllRow.push(r);
          }
          sumIndex=r+2;
        }
      }
      sheet.setColumnWidth(c, 80, GC.Spread.Sheets.SheetArea.rowHeader)
    }
    sumIndex=0
    let sumAllCol=[];//特殊处理列号
    //设置列标题
    for (var c = 0; c < getJsonLength(colDismension); c++) {
      for (var r = 0; r < columnHeaderRowCount; r++) {
        let colName = Object.keys(colDismension[0])[r];
        for (let valIndex = 0; valIndex < valueHeaderRowCount; valIndex++) {
          sheet.setValue(r, c * valueHeaderRowCount + valIndex, colDismension[c][colName], GC.Spread.Sheets.SheetArea.colHeader)
          //最后一列添加小计
          if(colDismension[c][colName]=='小计'){
            var currentColNum = c * valueHeaderRowCount + valIndex+1;
            for(let rowCal = 0; rowCal <= rowCount; rowCal++){
            let columnSpan = numToString(valIndex+sumIndex)+(rowCal+1)+":"+numToString(c * valueHeaderRowCount + valIndex - 1)+(rowCal+1);
            sheet.setFormula(rowCal, c * valueHeaderRowCount + valIndex, "SUMPRODUCT((MOD(COLUMN("+columnSpan+"),"+(columnHeaderRowCount-1)+")="+(valIndex+1)%valueHeaderRowCount+")*"+columnSpan+")");
          }
          sumAllCol.push(c * valueHeaderRowCount + valIndex);
        }
        else{
          currentColNum=sumIndex;
        }
      }
      sumIndex=currentColNum;
      }
      sheet.setRowHeight(r, 30, GC.Spread.Sheets.SheetArea.colHeader)
    }

    //行合计
    if(gridSetting.rowSum){
      sheet.addRows(rowCount,1,GC.Spread.Sheets.SheetArea.viewport);
      sheet.addSpan(rowCount, 0,1,rowHeadersColumnCount,GC.Spread.Sheets.SheetArea.rowHeader);
      sheet.setValue(rowCount, 0, "总计",GC.Spread.Sheets.SheetArea.rowHeader);
      for(let colCal = 0; colCal <= columnCount; colCal++){
            let formula='';
            if(sumAllCol.indexOf(colCal)<0){
              formula="SUBTOTAL(109,"+numToString(colCal)+  1 + ":"+numToString(colCal) + rowCount + ")"
            }
            else{
              formula="SUBTOTAL(109,"+numToString(colCal)+  1 + ":"+numToString(colCal) + rowCount + ")/2"
            }
            sheet.setFormula(rowCount, colCal, formula);
      }
    }
    
    //列合计
    if(gridSetting.colSum){
      sheet.addColumns(columnCount,valueHeaderRowCount,GC.Spread.Sheets.SheetArea.viewport);
      sheet.addSpan(0, columnCount,columnHeaderRowCount-1,valueHeaderRowCount,GC.Spread.Sheets.SheetArea.colHeader);
      sheet.setValue(0, columnCount , "总计",GC.Spread.Sheets.SheetArea.colHeader);
      for (var valIndex = 0; valIndex < valueHeaderRowCount; valIndex++) {
        let ix = valIndex % valueHeaderRowCount;
        //alert(valDismension[ix]);
        sheet.setValue(columnHeaderRowCount - 1,columnCount+valIndex, valDismension[ix], GC.Spread.Sheets.SheetArea.colHeader)
        for(let rowCal = 0; rowCal <= rowCount; rowCal++){
          let columnSpan = numToString(valIndex)+(rowCal+1)+":"+numToString(columnCount-valueHeaderRowCount+valIndex)+(rowCal+1);
          let formula='';
          if(sumAllRow.length>0){
              formula="SUMPRODUCT((MOD(COLUMN("+columnSpan+"),"+(valueHeaderRowCount)+")="+(valIndex+1)%valueHeaderRowCount+")*"+columnSpan+")/2";
            }
          sheet.setFormula(rowCal, columnCount+valIndex, formula);
        }
      }
    }
    
    //设置值列标题
    for (var valIndex = 0; valIndex < columnCount; valIndex++) {
      let ix = valIndex % valueHeaderRowCount;
      sheet.setValue(columnHeaderRowCount - 1, valIndex, valDismension[ix], GC.Spread.Sheets.SheetArea.colHeader)
    }
    //设置数据
    for (let data of datas) {
      //查找行坐标
      let xAis = -1;
      let valueValid = true;
      for (let row of rowDismension) {
        xAis++;
        for (let a = 0; a < rowHeadersColumnCount; a++) {
          let key = Object.keys(rowDismension[0])[a];
          if (data[key] != row[key]) {
            valueValid = false
            break;
          }
          else {
            valueValid = true;
          }
        }
        if (!valueValid) continue
        //查找列坐标
        let yAis = -1;
        for (let col of colDismension) {
          yAis++;
          for (let a = 0; a < columnHeaderRowCount - 1; a++) {
            let key = Object.keys(colDismension[0])[a];
            if (data[key] != col[key]) {
              valueValid = false
              break;
            }
            else {
              valueValid = true;
            }
          }
          if (!valueValid) continue

          //表格填充值
          for (let a = 0; a < valueHeaderRowCount; a++) {
            let key = valDismension[a];
            sheet.setValue(xAis, yAis * 2 + a, data[key], GC.Spread.Sheets.SheetArea.viewport);
          }
          break;
        }
        break;
      }
    }
    //设置整体颜色
    for (var r = 0; r < rowCount; r++) {
      let rowLevel = 1;
      if (r%2 == 1) {
        sheet.getRange(r, -1, 1, -1).backColor('rgba(0,133,199,.15)')
        mergeRow(sheet, r, 0, rowHeadersColumnCount, GC.Spread.Sheets.SheetArea.rowHeader);
      }
    }
    for (var c = 0; c < columnCount; c++) {
      sheet.setColumnWidth(c, 100);
      if (c%2 == -1) {
        sheet.getRange(-1, c, -1, 1).backColor('rgba(0,133,199,.15)')
        mergeColumn(sheet, 0, c, columnHeaderRowCount, GC.Spread.Sheets.SheetArea.colHeader);
      }
    }
    //合并行标题
    for (var r = 0; r < columnHeaderRowCount; r++) {
      if (r == 0) {
        mergeRow(sheet, r, 0, columnCount, GC.Spread.Sheets.SheetArea.colHeader);
      }
      else {
        let c = 0;
        while (c < columnCount) {
          let span = sheet.getSpan(r - 1, c, GC.Spread.Sheets.SheetArea.colHeader)
          if (span && span.colCount > 1) {
            mergeRow(sheet, r, c, span.colCount, GC.Spread.Sheets.SheetArea.colHeader);
            c += span.colCount;
          }
          else {
            c++;
          }
        }
      }
    }
    //合并列标题
    for (let c = 0; c < rowHeadersColumnCount; c++) {
      if (c == 0) {
        mergeColumn(sheet, 0, c, rowCount, GC.Spread.Sheets.SheetArea.rowHeader);
      }
      else {
        let r = 0;
        while (r < rowCount) {
          let span = sheet.getSpan(r, c - 1, GC.Spread.Sheets.SheetArea.rowHeader)
          if (span && span.rowCount > 1) {
            mergeColumn(sheet, r, c, span.rowCount, GC.Spread.Sheets.SheetArea.rowHeader);
            r += span.rowCount;
          }
          else {
            r++;
          }
        }
      }
    }

    sheet.resumePaint();
  }
}
function mergeColumn(sheet:any, row:any, col:any, rowCount:any, sheetArea:any) {
  if (rowCount < 2) {
    return;
  }
  try {
    let c = col, startRow = row, rowLength = 1;
    let oldValue = sheet.getValue(row, c, sheetArea);
    for (let r = row + 1; r < row + rowCount; r++) {
      let value = sheet.getValue(r, c, sheetArea);
      if (oldValue === value) {
        rowLength++;
      }
      else {
        if (rowLength > 1) {
          sheet.addSpan(startRow, c, rowLength, 1, sheetArea)
        }
        oldValue = value;
        startRow = r;
        rowLength = 1;
      }
    }
    if (rowLength > 1) {
      sheet.addSpan(startRow, c, rowLength, 1, sheetArea)
    }
  }
  catch (e) {

  }
}

function mergeRow(sheet:any, row:any, col:any, columnCount:any, sheetArea:any) {
  if (columnCount < 2) {
    return;
  }
  try {
    let r = row, startCol = col, colLength = 1;
    let oldValue = sheet.getValue(r, col, sheetArea);
    for (let c = col + 1; c < col + columnCount; c++) {
      let value = sheet.getValue(r, c, sheetArea);
      if (oldValue === value) {
        colLength++;
      }
      else {
        if (colLength > 1) {
          sheet.addSpan(r, startCol, 1, colLength, sheetArea);
        }
        oldValue = value;
        startCol = c;
        colLength = 1;
      }
    }
    if (colLength > 1) {
      sheet.addSpan(r, startCol, 1, colLength, sheetArea)
    }
  }
  catch (e) {

  }
}

function resetActiveSheet(spread:any, sheet:any) {
  let name = sheet.name();
  let index = spread.getSheetIndex(name);
  spread.removeSheet(index)
  spread.addSheet(index, new GC.Spread.Sheets.Worksheet(name));
}

// gets a random integer between zero and max
function randomInt(max:any) {
  return Math.floor(Math.random() * (max + 1));
}

// gets a simple data set for basic demos
function getSimpleDataSet(cnt:any) {
  let data = [
      { row1: 'Wijmo', row2: 'USA', value1: 2, value2: 222, col1: '2019-01-30', col2: '正确' },
      { row1: 'Wijmo', row2: 'Japan', value1: 1, value2: 2, col1: '2019-01-30', col2: '正确' },
      { row1: 'Aoba', row2: 'USA', value1: 1, value2: 200, col1: '2019-01-30', col2: '正确' },
      { row1: 'Aoba', row2: 'Japan', value1: 1, value2: 200, col1: '2019-01-31', col2: '正确' },
      { row1: 'Aoba', row2: 'Chinese', value1: 10, value2: 100, col1: '2019-01-31', col2: '正确' },
      { row1: 'Wijmo', row2: '上海', value1: 222, value2: 2222, col1: '2019-01-30', col2: '正确' },
      { row1: 'Aoba', row2: '北京', value1: 10, value2: 100, col1: '2019-01-30', col2: '正确' },
      { row1: 'Wijmo', row2: '上海', value1: 10, value2: 100, col1: '2019-01-31', col2: '错误' },
      { row1: 'Aoba', row2: '北京', value1: 10, value2: 100, col1: '2019-01-31', col2: '错误' }
  ];
  return data;
}
function getJsonLength(jsonData: any) {
  let jsonLength = 0;
  for (let item in jsonData) {
    jsonLength++;
  }
  return jsonLength;
}
function JsonQuery(arr: Array<any>, obj: any): any{
  let _arr: any = [];
  for (let _jsonObj of arr) {
    let _b: boolean = true;
    for (let prop in obj) {
      if (_jsonObj[prop] != obj[prop]) {
        _b = false;
        break;
      }
    }
    if (_b) _arr.push(_jsonObj)
  }
  return _arr;
}
function numToString(numm:any){
    let stringArray:any [] = [];
    stringArray.length = 0;
    let numToStringAction = function(nnum:any){
        let num = nnum;
        let a = parseInt((num / 26).toString());
        let b = parseInt((num % 26).toString());
        stringArray.push(String.fromCharCode(64 +(b+1)));
        if(a>0){
            numToStringAction(a);
        }
    }
    numToStringAction(numm);
    return stringArray.reverse().join("");
}
function refreshList(slicerData:any) {
    var filteredRowIndexs = slicerData.getFilteredRowIndexes();
    /*
    var trs = document.getElementsByTagName('tr');

    for (var i = 0; i < slicerData.data.length; i++) {
        if (filteredRowIndexs.indexOf(i) !== -1) {
            trs[i + 1].style.display = '';
        } else {
            trs[i + 1].style.display = 'none';
        }
    }
    */
}