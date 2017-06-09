---
title: LegendKey.Creator Property (Excel)
keywords: vbaxl10.chm589074
f1_keywords:
- vbaxl10.chm589074
ms.prod: excel
api_name:
- Excel.LegendKey.Creator
ms.assetid: de496f53-4edc-509a-7d5e-a2a9b28b25a2
ms.date: 06/08/2017
---


# LegendKey.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **LegendKey** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[LegendKey Object](legendkey-object-excel.md)

