---
title: Axes.Application Property (Excel)
keywords: vbaxl10.chm571073
f1_keywords:
- vbaxl10.chm571073
ms.prod: excel
api_name:
- Excel.Axes.Application
ms.assetid: 69b31571-68ad-dfb8-ea28-529cfa150132
ms.date: 06/08/2017
---


# Axes.Application Property (Excel)

When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.


## Syntax

 _expression_ . **Application**

 _expression_ A variable that represents an **Axes** object.


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


[Axes Collection](axes-object-excel.md)

