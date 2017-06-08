---
title: ListObject.HeaderRowRange Property (Excel)
keywords: vbaxl10.chm734084
f1_keywords:
- vbaxl10.chm734084
ms.prod: excel
api_name:
- Excel.ListObject.HeaderRowRange
ms.assetid: af7ca1d5-f72f-f369-9946-c64eb0cf9da0
ms.date: 06/08/2017
---


# ListObject.HeaderRowRange Property (Excel)

Returns a  **[Range](range-object-excel.md)** object that represents the range of the header row for a list. Read-only **Range** .


## Syntax

 _expression_ . **HeaderRowRange**

 _expression_ A variable that represents a **ListObject** object.


## Example

The following example activates the range specified by the  **HeaderRowRange** property of the default **[ListObject](listobject-object-excel.md)** object in the first worksheet of the active workbook.


```vb
Sub ActivateHeaderRow() 
 Dim wrksht As Worksheet 
 Dim objList As ListObject 
 Dim objListRng As Range 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set objList = wrksht.ListObjects(1) 
 Set objListRng = objList.HeaderRowRange 
 
 objListRng.Activate 
End Sub
```


## See also


#### Concepts


[ListObject Object](listobject-object-excel.md)

