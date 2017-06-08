---
title: ListRows.Application Property (Excel)
keywords: vbaxl10.chm739073
f1_keywords:
- vbaxl10.chm739073
ms.prod: excel
api_name:
- Excel.ListRows.Application
ms.assetid: 556e3016-4cfb-9e15-a2b4-7fc651e10859
ms.date: 06/08/2017
---


# ListRows.Application Property (Excel)

When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.


## Syntax

 _expression_ . **Application**

 _expression_ A variable that represents a **ListRows** object.


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


[ListRows Object](listrows-object-excel.md)

