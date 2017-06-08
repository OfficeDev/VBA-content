---
title: FullSeriesCollection.Application Property (Excel)
keywords: vbaxl10.chm943073
f1_keywords:
- vbaxl10.chm943073
ms.prod: excel
ms.assetid: 52dfb5aa-c6fb-201c-c1ed-880aff1efb45
ms.date: 06/08/2017
---


# FullSeriesCollection.Application Property (Excel)

Returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. Read-only.


## Syntax

 _expression_ . **Application**

 _expression_ A variable that represents a[FullSeriesCollection Object (Excel)](fullseriescollection-object-excel.md) object.


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



[FullSeriesCollection Object](fullseriescollection-object-excel.md)

