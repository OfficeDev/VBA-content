---
title: ModelFormatDate.Creator Property (Excel)
keywords: vbaxl10.chm983074
f1_keywords:
- vbaxl10.chm983074
ms.assetid: 4f7b44a5-70da-be7d-306c-9a2d2c9ea724
ms.date: 06/08/2017
ms.prod: excel
---


# ModelFormatDate.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **ModelFormatDate** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The Creator  **property** is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Other resources


[ModelFormatDate Object](modelformatdate-object-excel.md)


