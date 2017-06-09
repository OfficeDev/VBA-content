---
title: Worksheet.ProtectContents Property (Excel)
keywords: vbaxl10.chm174090
f1_keywords:
- vbaxl10.chm174090
ms.prod: excel
api_name:
- Excel.Worksheet.ProtectContents
ms.assetid: 807717f6-1265-2d5d-5221-bc46b24d8281
ms.date: 06/08/2017
---


# Worksheet.ProtectContents Property (Excel)

 **True** if the contents of the sheet are protected. This protects the individual cells. To turn on content protection, use the **[Protect](worksheet-protect-method-excel.md)** method with the _Contents_ argument set to **True** . Read-only **Boolean** .


## Syntax

 _expression_ . **ProtectContents**

 _expression_ A variable that represents a **Worksheet** object.


## Example

This example displays a message box if the contents of Sheet1 are protected.


```vb
If Worksheets("Sheet1").ProtectContents = True Then 
 MsgBox "The contents of Sheet1 are protected." 
End If
```


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

