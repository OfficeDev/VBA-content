---
title: PivotFilter.Creator Property (Excel)
keywords: vbaxl10.chm769074
f1_keywords:
- vbaxl10.chm769074
ms.prod: excel
api_name:
- Excel.PivotFilter.Creator
ms.assetid: b35b64b0-2565-cd94-92c0-a013e3fbb4a5
ms.date: 06/08/2017
---


# PivotFilter.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **PivotFilter** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[PivotFilter Object](pivotfilter-object-excel.md)

