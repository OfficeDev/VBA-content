---
title: PivotAxis.Creator Property (Excel)
keywords: vbaxl10.chm767074
f1_keywords:
- vbaxl10.chm767074
ms.prod: excel
api_name:
- Excel.PivotAxis.Creator
ms.assetid: 4fa167dd-6cc3-f296-7d34-15dc835d7310
ms.date: 06/08/2017
---


# PivotAxis.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **PivotAxis** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[PivotAxis Object](pivotaxis-object-excel.md)

