---
title: PivotTables.Add Method (Excel)
keywords: vbaxl10.chm238076
f1_keywords:
- vbaxl10.chm238076
ms.prod: excel
api_name:
- Excel.PivotTables.Add
ms.assetid: 3b830532-e834-81c8-dd5e-a43ed2efc269
ms.date: 06/08/2017
---


# PivotTables.Add Method (Excel)

Adds a new PivotTable report. Returns a  **[PivotTable](pivottable-object-excel.md)** object.


## Syntax

 _expression_ . **Add**( **_PivotCache_** , **_TableDestination_** , **_TableName_** , **_ReadData_** , **_DefaultVersion_** )

 _expression_ A variable that represents a **PivotTables** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _PivotCache_|Required| **[PivotCache](pivotcache-object-excel.md)**|The PivotTable cache on which the new PivotTable report is based. The cache provides data for the report.|
| _TableDestination_|Required| **Variant**|The cell in the upper-left corner of the PivotTable report's destination range (the range on the worksheet where the resulting report will be placed). You must specify a destination range on the worksheet that contains the  **PivotTables** object specified by _expression_.|
| _TableName_|Optional| **Variant**|The name of the new PivotTable report.|
| _ReadData_|Optional| **Variant**| **True** to create a PivotTable cache that contains all records from the external database; this cache can be very large. **False** to enable setting some of the fields as server-based page fields before the data is actually read.|
| _DefaultVersion_|Optional| **Variant**|The version of Microsoft Excel the PivotTable was originally created in.|

### Return Value

A  **[PivotTable](pivottable-object-excel.md)** object that represents the new PivotTable report.


## Example

This example creates a new PivotTable cache based on an OLAP provider, and then it creates a new PivotTable report based on the cache, at cell A1 on the first worksheet.


```vb
Dim cnnConn As ADODB.Connection 
Dim rstRecordset As ADODB.Recordset 
Dim cmdCommand As ADODB.Command 
 
' Open the connection. 
Set cnnConn = New ADODB.Connection 
With cnnConn 
 .ConnectionString = _ 
 "Provider=Microsoft.Jet.OLEDB.4.0" 
 .Open "C:\perfdate\record.mdb" 
End With 
 
' Set the command text. 
Set cmdCommand = New ADODB.Command 
Set cmdCommand.ActiveConnection = cnnConn 
With cmdCommand 
 .CommandText = "Select Speed, Pressure, Time From DynoRun" 
 .CommandType = adCmdText 
 .Execute 
End With 
 
' Open the recordset. 
Set rstRecordset = New ADODB.Recordset 
Set rstRecordset.ActiveConnection = cnnConn 
rstRecordset.Open cmdCommand 
 
' Create PivotTable cache and report. 
Set objPivotCache = ActiveWorkbook.PivotCaches.Add( _ 
 SourceType:=xlExternal) 
Set objPivotCache.Recordset = rstRecordset 
 
ActiveSheet.PivotTables.Add _ 
 PivotCache:=objPivotCache, _ 
 TableDestination:=Range("A3"), _ 
 TableName:="Performance" 
 
With ActiveSheet.PivotTables("Performance") 
 .SmallGrid = False 
 With .PivotFields("Pressure") 
 .Orientation = xlRowField 
 .Position = 1 
 End With 
 With .PivotFields("Speed") 
 .Orientation = xlColumnField 
 .Position = 1 
 End With 
 With .PivotFields("Time") 
 .Orientation = xlDataField 
 .Position = 1 
 End With 
End With 
 
' Close the connections and clean up. 
cnnConn.Close 
Set cmdCommand = Nothing 
Set rstRecordSet = Nothing 
Set cnnConn = Nothing
```


## See also


#### Concepts


[PivotTables Object](pivottables-object-excel.md)

