---
title: FormatCondition.Creator Property (Excel)
keywords: vbaxl10.chm511074
f1_keywords:
- vbaxl10.chm511074
ms.prod: excel
api_name:
- Excel.FormatCondition.Creator
ms.assetid: f089db52-af38-22a4-7475-9803c64b9722
ms.date: 06/08/2017
---


# FormatCondition.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **FormatCondition** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[FormatCondition Object](formatcondition-object-excel.md)

