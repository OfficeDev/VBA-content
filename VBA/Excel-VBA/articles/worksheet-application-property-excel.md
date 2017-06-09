---
title: Worksheet.Application Property (Excel)
keywords: vbaxl10.chm173073
f1_keywords:
- vbaxl10.chm173073
ms.prod: excel
api_name:
- Excel.Worksheet.Application
ms.assetid: 65ca3e7e-2b8f-5882-baaa-17b7658bbf8b
ms.date: 06/08/2017
---


# Worksheet.Application Property (Excel)

When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.


## Syntax

 _expression_ . **Application**

 _expression_ A variable that represents a **Worksheet** object.


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


[Worksheet Object](worksheet-object-excel.md)

