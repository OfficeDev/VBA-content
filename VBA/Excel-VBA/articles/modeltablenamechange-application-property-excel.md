---
title: ModelTableNameChange.Application Property (Excel)
keywords: vbaxl10.chm955073
f1_keywords:
- vbaxl10.chm955073
ms.prod: excel
ms.assetid: d393786f-8f33-ed78-42a3-436e92c2b704
ms.date: 06/08/2017
---


# ModelTableNameChange.Application Property (Excel)

Returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. Read-only.


## Syntax

 _expression_ . **Application**

 _expression_ A variable that represents a[ModelTableNameChange Object (Excel)](modeltablenamechange-object-excel.md) object.


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



[ModelTableNameChange Object](modeltablenamechange-object-excel.md)

