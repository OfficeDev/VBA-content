---
title: Range.Creator Property (Excel)
keywords: vbaxl10.chm143074
f1_keywords:
- vbaxl10.chm143074
ms.prod: excel
api_name:
- Excel.Range.Creator
ms.assetid: d7970f19-b10d-9101-4326-ea2d2460e849
ms.date: 06/08/2017
---


# Range.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **Range** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[Range Object](range-object-excel.md)

