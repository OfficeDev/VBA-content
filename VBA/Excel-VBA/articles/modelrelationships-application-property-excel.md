---
title: ModelRelationships.Application Property (Excel)
keywords: vbaxl10.chm939073
f1_keywords:
- vbaxl10.chm939073
ms.prod: excel
ms.assetid: 8c2d631a-84bc-8709-79ba-bffe40ed676f
ms.date: 06/08/2017
---


# ModelRelationships.Application Property (Excel)

Returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. Read-only.


## Syntax

 _expression_ . **Application**

 _expression_ A variable that represents a[ModelRelationships Object (Excel)](modelrelationships-object-excel.md) object.


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



[ModelRelationships Object](modelrelationships-object-excel.md)

