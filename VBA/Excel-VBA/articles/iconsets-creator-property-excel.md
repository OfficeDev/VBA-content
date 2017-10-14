---
title: IconSets.Creator Property (Excel)
keywords: vbaxl10.chm819074
f1_keywords:
- vbaxl10.chm819074
ms.prod: excel
api_name:
- Excel.IconSets.Creator
ms.assetid: e46acfe1-71f0-3a10-92d9-dd1ab3aa5569
ms.date: 06/08/2017
---


# IconSets.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents an **IconSets** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[IconSets Object](iconsets-object-excel.md)

