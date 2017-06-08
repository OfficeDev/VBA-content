---
title: PivotTableChangeList.Creator Property (Excel)
keywords: vbaxl10.chm890074
f1_keywords:
- vbaxl10.chm890074
ms.prod: excel
api_name:
- Excel.PivotTableChangeList.Creator
ms.assetid: e843c050-3fe0-8aaa-85e3-7ca3b925ba8d
ms.date: 06/08/2017
---


# PivotTableChangeList.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **[PivotTableChangeList](pivottablechangelist-object-excel.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[PivotTableChangeList Object](pivottablechangelist-object-excel.md)

