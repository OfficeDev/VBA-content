---
title: ModelColumnName.Application Property (Excel)
keywords: vbaxl10.chm961073
f1_keywords:
- vbaxl10.chm961073
ms.prod: excel
ms.assetid: a15b21c5-0d29-8e5c-2d85-0d8d5810fba1
ms.date: 06/08/2017
---


# ModelColumnName.Application Property (Excel)

Returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. Read-only.


## Syntax

 _expression_ . **Application**

 _expression_ A variable that represents a[ModelColumnName Object (Excel)](modelcolumnname-object-excel.md) object.


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



[ModelColumnName Object](modelcolumnname-object-excel.md)

