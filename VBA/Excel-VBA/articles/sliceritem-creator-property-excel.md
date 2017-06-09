---
title: SlicerItem.Creator Property (Excel)
keywords: vbaxl10.chm906074
f1_keywords:
- vbaxl10.chm906074
ms.prod: excel
api_name:
- Excel.SlicerItem.Creator
ms.assetid: 66027cd8-f471-c194-9d3e-b19198e1cc2d
ms.date: 06/08/2017
---


# SlicerItem.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **[SlicerItem](sliceritem-object-excel.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[SlicerItem Object](sliceritem-object-excel.md)

