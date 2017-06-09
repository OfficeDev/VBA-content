---
title: ModelFormatWholeNumber.Creator Property (Excel)
keywords: vbaxl10.chm987074
f1_keywords:
- vbaxl10.chm987074
ms.assetid: 82f16ccb-6f50-273e-5ed4-e16db1262ecc
ms.date: 06/08/2017
ms.prod: excel
---


# ModelFormatWholeNumber.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **ModelFormatWholeNumber** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The Creator  **property** is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Other resources


[ModelFormatWholeNumber Object](modelformatwholenumber-object-excel.md)


