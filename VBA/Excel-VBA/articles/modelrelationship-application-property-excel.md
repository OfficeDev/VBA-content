---
title: ModelRelationship.Application Property (Excel)
keywords: vbaxl10.chm937073
f1_keywords:
- vbaxl10.chm937073
ms.prod: excel
ms.assetid: fc6832ad-4100-e1ac-f286-6f0cbe11c983
ms.date: 06/08/2017
---


# ModelRelationship.Application Property (Excel)

Returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. Read-only.


## Syntax

 _expression_ . **Application**

 _expression_ A variable that represents a[ModelRelationship Object (Excel)](modelrelationship-object-excel.md) object.


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



[ModelRelationship Object](modelrelationship-object-excel.md)

