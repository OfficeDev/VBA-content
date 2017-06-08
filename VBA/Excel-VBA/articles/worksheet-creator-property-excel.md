---
title: Worksheet.Creator Property (Excel)
keywords: vbaxl10.chm173074
f1_keywords:
- vbaxl10.chm173074
ms.prod: excel
api_name:
- Excel.Worksheet.Creator
ms.assetid: 39bb2896-2a2f-a7b2-8139-40f0f37104ed
ms.date: 06/08/2017
---


# Worksheet.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **Worksheet** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

