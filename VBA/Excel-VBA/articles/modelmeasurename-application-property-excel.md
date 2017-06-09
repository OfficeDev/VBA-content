---
title: ModelMeasureName.Application Property (Excel)
keywords: vbaxl10.chm969073
f1_keywords:
- vbaxl10.chm969073
ms.prod: excel
ms.assetid: 2a93826c-7d6d-030c-e0e3-1c9b85be9c4c
ms.date: 06/08/2017
---


# ModelMeasureName.Application Property (Excel)

Returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. Read-only.


## Syntax

 _expression_ . **Application**

 _expression_ A variable that represents a[ModelMeasureName Object (Excel)](modelmeasurename-object-excel.md) object.


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



[ModelMeasureName Object](modelmeasurename-object-excel.md)

