---
title: QuickAnalysis.Application Property (Excel)
keywords: vbaxl10.chm919073
f1_keywords:
- vbaxl10.chm919073
ms.prod: excel
ms.assetid: ad51f454-62a0-7eb7-b629-b72bd000e0e9
ms.date: 06/08/2017
---


# QuickAnalysis.Application Property (Excel)

Returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. Read-only.


## Syntax

 _expression_ . **Application**

 _expression_ A variable that represents a[QuickAnalysis Object (Excel)](quickanalysis-object-excel.md) object.


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



[QuickAnalysis Object](quickanalysis-object-excel.md)

