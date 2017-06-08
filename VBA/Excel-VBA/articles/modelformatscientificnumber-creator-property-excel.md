---
title: ModelFormatScientificNumber.Creator Property (Excel)
keywords: vbaxl10.chm991074
f1_keywords:
- vbaxl10.chm991074
ms.assetid: b764b8cb-b6f4-dca8-9bab-6add833dc61b
ms.date: 06/08/2017
ms.prod: excel
---


# ModelFormatScientificNumber.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **ModelFormatScientificNumber** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The Creator  **property** is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Other resources


[ModelFormatScientificNumber Object](modelformatscientificnumber-object-excel.md)


