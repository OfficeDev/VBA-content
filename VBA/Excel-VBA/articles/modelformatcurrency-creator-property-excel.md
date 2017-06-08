---
title: ModelFormatCurrency.Creator Property (Excel)
keywords: vbaxl10.chm993074
f1_keywords:
- vbaxl10.chm993074
ms.assetid: 069eb7ee-2168-0820-1018-61c1498c7929
ms.date: 06/08/2017
ms.prod: excel
---


# ModelFormatCurrency.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **ModelFormatCurrency** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The Creator  **property** is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Other resources


[ModelFormatCurrency Object](modelformatcurrency-object-excel.md)


