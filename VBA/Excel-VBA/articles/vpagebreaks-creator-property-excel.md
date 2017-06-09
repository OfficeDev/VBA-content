---
title: VPageBreaks.Creator Property (Excel)
keywords: vbaxl10.chm166074
f1_keywords:
- vbaxl10.chm166074
ms.prod: excel
api_name:
- Excel.VPageBreaks.Creator
ms.assetid: d8ff8785-8cf5-de2f-0425-8a605a72e6da
ms.date: 06/08/2017
---


# VPageBreaks.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **VPageBreaks** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[VPageBreaks Object](vpagebreaks-object-excel.md)

