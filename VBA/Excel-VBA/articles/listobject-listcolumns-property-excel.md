---
title: ListObject.ListColumns Property (Excel)
keywords: vbaxl10.chm734086
f1_keywords:
- vbaxl10.chm734086
ms.prod: excel
api_name:
- Excel.ListObject.ListColumns
ms.assetid: 64cefe01-b0e6-1cdd-3eec-7cb8389666dc
ms.date: 06/08/2017
---


# ListObject.ListColumns Property (Excel)

Returns a  **[ListColumns](listcolumns-object-excel.md)** collection that represents all the columns in a **[ListObject](listobject-object-excel.md)** object. Read-only.


## Syntax

 _expression_ . **ListColumns**

 _expression_ A variable that represents a **ListObject** object.


## Example

The following example displays the name of the second column in the  **ListColumns** collection object as created by a call to the **ListColumns** property. For this code to run, the Sheet1 worksheet must contain a table.


```vb
Sub DisplayColumnName 
 Dim wrksht As Worksheet 
 Dim objListObj As ListObject 
 Dim objListCols As ListColumns 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set objListObj = wrksht.ListObjects(1) 
 Set objListCols = objListObj.ListColumns 
 
 Debug.Print objListCols(2).Name 
End Sub
```


## See also


#### Concepts


[ListObject Object](listobject-object-excel.md)

