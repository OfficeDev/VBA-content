---
title: Style.NameLocal Property (Excel)
keywords: vbaxl10.chm177091
f1_keywords:
- vbaxl10.chm177091
ms.prod: excel
api_name:
- Excel.Style.NameLocal
ms.assetid: fcc978b3-c23b-8a5f-9e5b-e815ecb2f92e
ms.date: 06/08/2017
---


# Style.NameLocal Property (Excel)

Returns or sets the name of the object, in the language of the user. Read-only  **String** .


## Syntax

 _expression_ . **NameLocal**

 _expression_ A variable that represents a **Style** object.


## Remarks

If the style is a built-in style, this property returns the name of the style in the language of the current locale.


## Example

This example displays the name and localized name of style one in the active workbook.


```vb
With ActiveWorkbook.Styles(1) 
 MsgBox "The name of the style is " &; .Name 
 MsgBox "The localized name of the style is " &; .NameLocal 
End With
```


## See also


#### Concepts


[Style Object](style-object-excel.md)

