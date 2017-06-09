---
title: SlicerPivotTables.Creator Property (Excel)
keywords: vbaxl10.chm910074
f1_keywords:
- vbaxl10.chm910074
ms.prod: excel
api_name:
- Excel.SlicerPivotTables.Creator
ms.assetid: 7c1bf1f9-4d70-4b21-b235-d0f89b2bd500
ms.date: 06/08/2017
---


# SlicerPivotTables.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **[SlicerPivotTables](slicerpivottables-object-excel.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[SlicerPivotTables Object](slicerpivottables-object-excel.md)

