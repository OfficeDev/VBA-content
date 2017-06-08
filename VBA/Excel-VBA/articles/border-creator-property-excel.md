---
title: Border.Creator Property (Excel)
keywords: vbaxl10.chm546074
f1_keywords:
- vbaxl10.chm546074
ms.prod: excel
api_name:
- Excel.Border.Creator
ms.assetid: 3135c4a4-fab8-6d7f-85da-909a290c1b1e
ms.date: 06/08/2017
---


# Border.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **Border** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[Border Object](border-object-excel.md)

