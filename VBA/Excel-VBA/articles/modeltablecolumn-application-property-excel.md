---
title: ModelTableColumn.Application Property (Excel)
keywords: vbaxl10.chm929073
f1_keywords:
- vbaxl10.chm929073
ms.prod: excel
ms.assetid: 69540e35-6a9a-0fd9-23b1-31457b33ba68
ms.date: 06/08/2017
---


# ModelTableColumn.Application Property (Excel)

Returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. Read-only.


## Syntax

 _expression_ . **Application**

 _expression_ A variable that represents a[ModelTableColumn Object (Excel)](modeltablecolumn-object-excel.md) object.


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



[ModelTableColumn Object](modeltablecolumn-object-excel.md)

