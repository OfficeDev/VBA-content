---
title: ListObject.InsertRowRange Property (Excel)
keywords: vbaxl10.chm734085
f1_keywords:
- vbaxl10.chm734085
ms.prod: excel
api_name:
- Excel.ListObject.InsertRowRange
ms.assetid: 5957f802-96b8-60a3-74e4-d7abcea7544b
ms.date: 06/08/2017
---


# ListObject.InsertRowRange Property (Excel)

 Returns a **Range** object representing the Insert row, if any, of a specified **[ListObject](listobject-object-excel.md)** object. Read-only **Range** .


## Syntax

 _expression_ . **InsertRowRange**

 _expression_ A variable that represents a **ListObject** object.


## Remarks

If the Insert row is not visible because the list is inactive, the  **Nothing** object will be returned.


## Example

The following example activates the range specified by the  **InsertRowRange** in the default **ListObject** object in the first worksheet of the active workbook.


```vb
Function ActivateInsertRow() As Boolean 
 
 Dim wrksht As Worksheet 
 Dim objList As ListObject 
 Dim objListRng As Range 
 
 Set wrksht = ActiveWorkbook.Worksheets(1) 
 Set objList = wrksht.ListObjects(1) 
 Set objListRng = objList.InsertRowRange 
 
 If objListRng Is Nothing Then 
 ActivateInsertRow = False 
 Else 
 objListRng.Activate 
 ActivateInsertRow = True 
 End If 
 
End Function
```


## See also


#### Concepts


[ListObject Object](listobject-object-excel.md)

