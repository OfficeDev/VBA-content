---
title: Style.Name Property (Excel)
keywords: vbaxl10.chm177090
f1_keywords:
- vbaxl10.chm177090
ms.prod: excel
api_name:
- Excel.Style.Name
ms.assetid: 4ad63465-afe0-fc96-3ec7-62318d8ac1e2
ms.date: 06/08/2017
---


# Style.Name Property (Excel)

Returns a  **String** value that represents the name of the object.


## Syntax

 _expression_ . **Name**

 _expression_ A variable that represents a **Style** object.


## Example

This example displays the name of style one in the active workbook, first in the language of the macro and then in the language of the user.


```vb
With ActiveWorkbook.Styles(1) 
 MsgBox "The name of the style: " &; .Name 
 MsgBox "The localized name of the style: " &; .NameLocal 
End With
```

The following example displays the name of the default  **ListObject** object in sheet1 of the active workbook.




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


[Style Object](style-object-excel.md)

