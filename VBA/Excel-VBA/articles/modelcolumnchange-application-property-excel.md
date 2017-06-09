---
title: ModelColumnChange.Application Property (Excel)
keywords: vbaxl10.chm965073
f1_keywords:
- vbaxl10.chm965073
ms.prod: excel
ms.assetid: 42065d25-aaef-e92a-f174-47f056e1e460
ms.date: 06/08/2017
---


# ModelColumnChange.Application Property (Excel)

Returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. Read-only.


## Syntax

 _expression_ . **Application**

 _expression_ A variable that represents a[ModelColumnChange Object (Excel)](modelcolumnchange-object-excel.md) object.


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



[ModelColumnChange Object](modelcolumnchange-object-excel.md)

