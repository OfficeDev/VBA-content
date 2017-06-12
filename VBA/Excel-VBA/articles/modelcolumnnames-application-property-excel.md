---
title: ModelColumnNames.Application Property (Excel)
keywords: vbaxl10.chm963073
f1_keywords:
- vbaxl10.chm963073
ms.prod: excel
ms.assetid: 09a0a219-b4eb-4ead-f058-5b9a04e98dc9
ms.date: 06/08/2017
---


# ModelColumnNames.Application Property (Excel)

Returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. Read-only.


## Syntax

 _expression_ . **Application**

 _expression_ A variable that represents a[ModelColumnNames Object (Excel)](modelcolumnnames-object-excel.md) object.


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



[ModelColumnNames Object](modelcolumnnames-object-excel.md)

