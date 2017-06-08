---
title: ModelTableNameChanges.Application Property (Excel)
keywords: vbaxl10.chm957073
f1_keywords:
- vbaxl10.chm957073
ms.prod: excel
ms.assetid: c1c99f30-cfa7-206c-0353-41e0b8fca17a
ms.date: 06/08/2017
---


# ModelTableNameChanges.Application Property (Excel)

Returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. Read-only.


## Syntax

 _expression_ . **Application**

 _expression_ A variable that represents a[ModelTableNameChanges Object (Excel)](modeltablenamechanges-object-excel.md) object.


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



[ModelTableNameChanges Object](modeltablenamechanges-object-excel.md)

