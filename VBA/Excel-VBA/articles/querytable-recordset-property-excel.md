---
title: QueryTable.Recordset Property (Excel)
keywords: vbaxl10.chm518094
f1_keywords:
- vbaxl10.chm518094
ms.prod: excel
api_name:
- Excel.QueryTable.Recordset
ms.assetid: d9f4190e-c43c-5fe5-113d-18c8efcc2a27
ms.date: 06/08/2017
---


# QueryTable.Recordset Property (Excel)

Returns or sets a  **Recordset** object that's used as the data source for the specified query table. Read/write.


## Syntax

 _expression_ . **Recordset**

 _expression_ A variable that represents a **QueryTable** object.


## Remarks

If this property is used to overwrite an existing recordset, the change takes effect when the  **[Refresh](querytable-refresh-method-excel.md)** method is run.

If you import data using the user interface, data from a Web query or a text query is imported as a  **[QueryTable](querytable-object-excel.md)** object, while all other external data is imported as a **[ListObject](listobject-object-excel.md)** object.

If you import data using the object model, data from a Web query or a text query must be imported as a  **QueryTable** , while all other external data can be imported as either a **ListObject** or a **QueryTable** .

You can use the  **[QueryTable](listobject-querytable-property-excel.md)** property of the **ListObject** to access the **RecordSet** property.


## Example

This example changes the  **Recordset** object used with the first query table on the first worksheet and then refreshes the query table.


```vb
With Worksheets(1).QueryTables(1) 
 .Recordset = _ 
 Workbooks.OpenDatabase("c:\Nwind.mdb") _ 
 .OpenRecordset("employees") 
 .Refresh 
End With
```


## See also


#### Concepts


[QueryTable Object](querytable-object-excel.md)

