---
title: DataLabel.Creator Property (Excel)
keywords: vbaxl10.chm581074
f1_keywords:
- vbaxl10.chm581074
ms.prod: excel
api_name:
- Excel.DataLabel.Creator
ms.assetid: 9387a1d2-052a-3af1-dde9-ed8b3c4ce7d6
ms.date: 06/08/2017
---


# DataLabel.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **DataLabel** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[DataLabel Object](datalabel-object-excel.md)

