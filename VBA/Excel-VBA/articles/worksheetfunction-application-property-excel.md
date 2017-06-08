---
title: WorksheetFunction.Application Property (Excel)
keywords: vbaxl10.chm136073
f1_keywords:
- vbaxl10.chm136073
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Application
ms.assetid: 0cf8b842-cc2b-13b3-e9b7-ce0921566ca9
ms.date: 06/08/2017
---


# WorksheetFunction.Application Property (Excel)

When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object. You can use this property with an OLE Automation object to return the application of that object. Read-only.


## Syntax

 _expression_ . **Application**

 _expression_ A variable that represents a **WorksheetFunction** object.


## Example

This example displays a message about the application that created  `myObject`.


```vb
Set myObject = ActiveWorkbook 
If myObject.Application.Value = "Microsoft Excel" Then 
 MsgBox "This is an Excel Application object." 
Else 
 MsgBox "This is not an Excel Application object." 
End If
```


## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

