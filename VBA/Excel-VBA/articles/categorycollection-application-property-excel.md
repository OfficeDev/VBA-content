---
title: CategoryCollection.Application Property (Excel)
keywords: vbaxl10.chm947073
f1_keywords:
- vbaxl10.chm947073
ms.prod: excel
ms.assetid: cfae4e60-9cda-c43b-e1d5-78ba110dd21c
ms.date: 06/08/2017
---


# CategoryCollection.Application Property (Excel)

Returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. Read-only.


## Syntax

 _expression_ . **Application**

 _expression_ A variable that represents a[CategoryCollection Object (Excel)](categorycollection-object-excel.md) object.


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



[CategoryCollection Object](categorycollection-object-excel.md)

