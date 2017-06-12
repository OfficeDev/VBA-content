---
title: ModelMeasureNames.Application Property (Excel)
keywords: vbaxl10.chm971073
f1_keywords:
- vbaxl10.chm971073
ms.prod: excel
ms.assetid: c755709d-d0f0-ac56-8d57-39230fd92486
ms.date: 06/08/2017
---


# ModelMeasureNames.Application Property (Excel)

Returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. Read-only.


## Syntax

 _expression_ . **Application**

 _expression_ A variable that represents a[ModelMeasureNames Object (Excel)](modelmeasurenames-object-excel.md) object.


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



[ModelMeasureNames Object](modelmeasurenames-object-excel.md)

