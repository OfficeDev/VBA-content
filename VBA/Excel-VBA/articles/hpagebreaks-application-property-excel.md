---
title: HPageBreaks.Application Property (Excel)
keywords: vbaxl10.chm163073
f1_keywords:
- vbaxl10.chm163073
ms.prod: excel
api_name:
- Excel.HPageBreaks.Application
ms.assetid: a3516361-fda0-e2e3-f634-c70ee5e8f49c
ms.date: 06/08/2017
---


# HPageBreaks.Application Property (Excel)

When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.


## Syntax

 _expression_ . **Application**

 _expression_ A variable that represents a **HPageBreaks** object.


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


[HPageBreaks Object](hpagebreaks-object-excel.md)

