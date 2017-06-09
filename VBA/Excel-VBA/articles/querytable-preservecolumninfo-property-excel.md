---
title: QueryTable.PreserveColumnInfo Property (Excel)
keywords: vbaxl10.chm518110
f1_keywords:
- vbaxl10.chm518110
ms.prod: excel
api_name:
- Excel.QueryTable.PreserveColumnInfo
ms.assetid: 736b5b43-17f5-84ca-6e79-e9eca12fe077
ms.date: 06/08/2017
---


# QueryTable.PreserveColumnInfo Property (Excel)

 **True** if column sorting, filtering, and layout information is preserved whenever a query table is refreshed. The default value is **True** . Read/write **Boolean** .


## Syntax

 _expression_ . **PreserveColumnInfo**

 _expression_ A variable that represents a **QueryTable** object.


## Remarks

This property has an effect only when the query table is using a database connection.

You can set this property to  **False** for compatibility with earlier versions of Microsoft Excel.

If you import data using the user interface, data from a Web query or a text query is imported as a  **[QueryTable](querytable-object-excel.md)** object, while all other external data is imported as a **[ListObject](listobject-object-excel.md)** object.

If you import data using the object model, data from a Web query or a text query must be imported as a  **QueryTable** , while all other external data can be imported as either a **ListObject** or a **QueryTable** .

You can use the  **[QueryTable](listobject-querytable-property-excel.md)** property of the **ListObject** to access the **PreserveColumnInfo** property.


## Example

This example preserves column sorting, filtering, and layout information for compatibility with earlier versions of Microsoft Excel.


```vb
Dim cnnConnect As ADODB.Connection 
Dim rstRecordset As ADODB.Recordset 
 
Set cnnConnect = New ADODB.Connection 
cnnConnect.Open "Provider=SQLOLEDB;" &; _ 
 "Data Source=srvdata;" &; _ 
 "User ID=wadet;Password=4me2no;" 
 
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


## See also


#### Concepts


[QueryTable Object](querytable-object-excel.md)

