---
title: Filter.Creator Property (Excel)
keywords: vbaxl10.chm541074
f1_keywords:
- vbaxl10.chm541074
ms.prod: excel
api_name:
- Excel.Filter.Creator
ms.assetid: 648b0917-011b-ec4f-4a7a-7a56b070a8cd
ms.date: 06/08/2017
---


# Filter.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **Filter** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[Filter Object](filter-object-excel.md)

