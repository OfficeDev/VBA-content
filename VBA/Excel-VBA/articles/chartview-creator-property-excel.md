---
title: ChartView.Creator Property (Excel)
keywords: vbaxl10.chm780074
f1_keywords:
- vbaxl10.chm780074
ms.prod: excel
api_name:
- Excel.ChartView.Creator
ms.assetid: b79054ab-5f0b-c7a0-3247-6e6cfe0470cd
ms.date: 06/08/2017
---


# ChartView.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **ChartView** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[ChartView Object](chartview-object-excel.md)

