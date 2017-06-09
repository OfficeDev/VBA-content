---
title: ModelColumnChanges.Application Property (Excel)
keywords: vbaxl10.chm967073
f1_keywords:
- vbaxl10.chm967073
ms.prod: excel
ms.assetid: da204577-a5b9-41c5-8d54-997d839e0f48
ms.date: 06/08/2017
---


# ModelColumnChanges.Application Property (Excel)

Returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. Read-only.


## Syntax

 _expression_ . **Application**

 _expression_ A variable that represents a[ModelColumnChanges Object (Excel)](modelcolumnchanges-object-excel.md) object.


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


## Property value

 **APPLICATION**


## See also


#### Other resources



[ModelColumnChanges Object](modelcolumnchanges-object-excel.md)

