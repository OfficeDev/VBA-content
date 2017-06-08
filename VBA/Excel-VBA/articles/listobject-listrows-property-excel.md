---
title: ListObject.ListRows Property (Excel)
keywords: vbaxl10.chm734087
f1_keywords:
- vbaxl10.chm734087
ms.prod: excel
api_name:
- Excel.ListObject.ListRows
ms.assetid: 7b584f41-ffc0-abe4-e755-ef163bcbb2ed
ms.date: 06/08/2017
---


# ListObject.ListRows Property (Excel)

Returns a  **[ListRows](listrows-object-excel.md)** object that represents all the rows of data in the **[ListObject](listobject-object-excel.md)** object. Read-only.


## Syntax

 _expression_ . **ListRows**

 _expression_ A variable that represents a **ListObject** object.


## Remarks

The  **ListRows** object returned does not include the header, total, or Insert rows.


## Example

The following example deletes a row specified by number in the  **ListRows** collection that is created by a call to the **ListRows** property.


```vb
Sub DeleteListRow(iRowNumber As Integer) 
 Dim wrksht As Worksheet 
 Dim objListObj As ListObject 
 Dim objListRows As ListRows 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set objListObj = wrksht.ListObjects(1) 
 Set objListRows = objListObj.ListRows 
 
 If (iRowNumber <> 0) And (iRowNumber < objListRows.Count - 1) Then 
 objListRows(iRowNumber).Delete 
 End If 
End Sub
```


## See also


#### Concepts


[ListObject Object](listobject-object-excel.md)

