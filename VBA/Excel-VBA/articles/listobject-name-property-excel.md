---
title: ListObject.Name Property (Excel)
keywords: vbaxl10.chm734088
f1_keywords:
- vbaxl10.chm734088
ms.prod: excel
api_name:
- Excel.ListObject.Name
ms.assetid: fbbdf2f9-6c5f-6ebe-35b1-74aab63971a4
ms.date: 06/08/2017
---


# ListObject.Name Property (Excel)

Returns or sets a  **String** value that represents the name of the **[ListObject](listobject-object-excel.md)** object.


## Syntax

 _expression_ . **Name**

 _expression_ An expression that returns a **ListObject** object.


### Return Value

String


## Remarks

This name is used solely as a unique identifier for the  **[Item](listobjects-item-property-excel.md)** property of the **[ListObjects](listobjects-object-excel.md)** collection objects. This property can only be set through the object model.

By default, each  **ListObject** object name begins with the word "List", followed by a number (no spaces). If an attempt is made to set the **Name** property to a name already used by another **ListObject** object, a run-time error is thrown.


## Example

The following example displays the name of the default ListObject object in sheet1 of the active workbook.


```vb
Sub Test 
 Dim wrksht As Worksheet 
 Dim oListObj As ListObject 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set oListObj = wrksht.ListObjects(1) 
 
 MsgBox oListObj.Name 
End Sub
```


## See also


#### Concepts


[ListObject Object](listobject-object-excel.md)

