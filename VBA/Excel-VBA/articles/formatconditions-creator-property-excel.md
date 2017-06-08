---
title: FormatConditions.Creator Property (Excel)
keywords: vbaxl10.chm509074
f1_keywords:
- vbaxl10.chm509074
ms.prod: excel
api_name:
- Excel.FormatConditions.Creator
ms.assetid: c828685a-91a9-d70d-a8e6-33da541f1ae9
ms.date: 06/08/2017
---


# FormatConditions.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **FormatConditions** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[FormatConditions Object](formatconditions-object-excel.md)

