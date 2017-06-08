---
title: Model.Application Property (Excel)
keywords: vbaxl10.chm941073
f1_keywords:
- vbaxl10.chm941073
ms.prod: excel
ms.assetid: 201e79d9-8e9b-0ff4-5f6e-dcdd3911b08a
ms.date: 06/08/2017
---


# Model.Application Property (Excel)

Returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. Read-only.


## Syntax

 _expression_ . **Application**

 _expression_ A variable that represents a object.


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


[Model Object Members](http://msdn.microsoft.com/library/b2bd944a-3484-222b-b3d6-acd70a6ac28a%28Office.15%29.aspx)


