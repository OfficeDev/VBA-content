---
title: QueryTables.Add Method (Excel)
keywords: vbaxl10.chm521074
f1_keywords:
- vbaxl10.chm521074
ms.prod: excel
api_name:
- Excel.QueryTables.Add
ms.assetid: ac6cd03e-31aa-cd8c-aa67-a551894c6eb3
ms.date: 06/08/2017
---


# QueryTables.Add Method (Excel)

Creates a new query table.


## Syntax

 _expression_ . **Add**( **_Connection_** , **_Destination_** , **_Sql_** )

 _expression_ A variable that represents a **QueryTables** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Connection_|Required| **Variant**|The data source for the query table. Can be one of the following:<ul><li><p>A string containing an OLE DB or ODBC connection string. The ODBC connection string has the form "ODBC;<connection string>".</p></li><li><p>A <b>QueryTable</b>  object from which the query information is initially copied, including the connection string and the SQL text, but not including the <b>Destination</b>  range. Specifying a <b>QueryTable</b>  object causes the <b>Sql</b>  argument to be ignored.</p></li><li><p>An ADO or DAO <b>Recordset</b>  object. Data is read from the ADO or DAO recordset. Microsoft Excel retains the recordset until the query table is deleted or the connection is changed. The resulting query table cannot be edited.</p></li><li><p>A Web query. A string in the form "URL;<url>", where "URL;" is required but not localized and the rest of the string is used for the URL of the Web query.</p></li><li><p>Data Finder. A string in the form "FINDER;<data finder file path>" where "FINDER;" is required but not localized. The rest of the string is the path and file name of a Data Finder file (*.dqy or *.iqy). The file is read when the <b>Add</b>  method is run; subsequent calls to the <b><a href="querytable-connection-property-excel.md">Connection</a></b>  property of the query table will return strings beginning with "ODBC;" or "URL;" as appropriate.</p></li><li><p>A text file. A string in the form "TEXT;<text file path and name>", where TEXT is required but not localized.</p></li></ul>|
| _Destination_|Required| **Range**|The cell in the upper-left corner of the query table destination range (the range where the resulting query table will be placed). The destination range must be on the worksheet that contains the  **QueryTables** object specified by expression.|
| _Sql_|Optional| **Variant**|The SQL query string to be run on the ODBC data source. This argument is optional when you're using an ODBC data source (if you don't specify it here, you should set it by using the  **Sql** property of the query table before the table is refreshed). You cannot use this argument when a **QueryTable** object, text file, or ADO or DAO **Recordset** object is specified as the data source.|

### Return Value

A  **[QueryTable](querytable-object-excel.md)** object that represents the new query table.


## Remarks

A query created by this method isn't run until the  **[Refresh](querytable-refresh-method-excel.md)** method is called.


## Example

This example creates a query table based on an ADO recordset. The example preserves the existing column sorting and filtering settings and layout information for backward compatibility.


```vb
Dim cnnConnect As ADODB.Connection 
Dim rstRecordset As ADODB.Recordset 
 
Set cnnConnect = New ADODB.Connection 
cnnConnect.Open "Provider=SQLOLEDB;" &; _ 
    "Data Source=srvdata;" &; _ 
    "User ID=testac;Password=4me2no;" 
 
Set rstRecordset = New ADODB.Recordset 
rstRecordset.Open _ 
    Source:="Select Name, Quantity, Price From Products", _ 
    ActiveConnection:=cnnConnect, _ 
    CursorType:=adOpenDynamic, _ 
    LockType:=adLockReadOnly, _ 
    Options:=adCmdText 
 
With ActiveSheet.QueryTables.Add( _ 
        Connection:=rstRecordset, _ 
        Destination:=Range("A1")) 
    .Name = "Contact List" 
    .FieldNames = True 
    .RowNumbers = False 
    .FillAdjacentFormulas = False 
    .PreserveFormatting = True 
    .RefreshOnFileOpen = False 
    .BackgroundQuery = True 
    .RefreshStyle = xlInsertDeleteCells 
    .SavePassword = True 
    .SaveData = True 
    .AdjustColumnWidth = True 
    .RefreshPeriod = 0 
    .PreserveColumnInfo = True 
    .Refresh BackgroundQuery:=False 
End With
```

This example imports a fixed width text file into a new query table. The first column in the text file is five characters wide and is imported as text. The second column is four characters wide and is skipped. The remainder of the text file is imported into the third column and has the General format applied to it.




```vb
Set shFirstQtr = Workbooks(1).Worksheets(1) 
Set qtQtrResults = shFirstQtr.QueryTables.Add( _ 
    Connection := "TEXT;C:\My Documents\19980331.txt", 
    Destination := shFirstQtr.Cells(1,1)) 
With qtQtrResults 
    .TextFileParsingType = xlFixedWidth 
    .TextFileFixedColumnWidths := Array(5,4) 
    .TextFileColumnDataTypes := _ 
        Array(xlTextFormat, xlSkipColumn, xlGeneralFormat) 
    .Refresh 
End With
```

This example creates a new query table on the active worksheet.




```vb
sqlstring = "select 96Sales.totals from 96Sales where profit < 5" 
connstring = _ 
    "ODBC;DSN=96SalesData;UID=Rep21;PWD=NUyHwYQI;Database=96Sales" 
With ActiveSheet.QueryTables.Add(Connection:=connstring, _ 
        Destination:=Range("B1"), Sql:=sqlstring) 
    .Refresh 
End With
```


## See also


#### Concepts


[QueryTables Object](querytables-object-excel.md)

