---
title: PivotFormulas.Creator Property (Excel)
keywords: vbaxl10.chm232074
f1_keywords:
- vbaxl10.chm232074
ms.prod: excel
api_name:
- Excel.PivotFormulas.Creator
ms.assetid: 23be5a99-984e-1c8b-ceb3-17e101b442d5
ms.date: 06/08/2017
---


# PivotFormulas.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **PivotFormulas** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[PivotFormulas Object](pivotformulas-object-excel.md)

