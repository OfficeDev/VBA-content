---
title: Worksheet.ProtectDrawingObjects Property (Excel)
keywords: vbaxl10.chm174091
f1_keywords:
- vbaxl10.chm174091
ms.prod: excel
api_name:
- Excel.Worksheet.ProtectDrawingObjects
ms.assetid: a3733b3b-dca4-4131-e197-5c919d44c7bd
ms.date: 06/08/2017
---


# Worksheet.ProtectDrawingObjects Property (Excel)

 **True** if shapes are protected. To turn on shape protection, use the **[Protect](worksheet-protect-method-excel.md)** method with the _DrawingObjects_ argument set to **True** . Read-only **Boolean** .


## Syntax

 _expression_ . **ProtectDrawingObjects**

 _expression_ A variable that represents a **Worksheet** object.


## Example

This example displays a message box if the shapes on Sheet1 are protected.


```vb
If Worksheets("Sheet1").ProtectDrawingObjects = True Then 
 MsgBox "The shapes on Sheet1 are protected." 
End If
```


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

