---
title: ChartObjects.Creator Property (Excel)
keywords: vbaxl10.chm495074
f1_keywords:
- vbaxl10.chm495074
ms.prod: excel
api_name:
- Excel.ChartObjects.Creator
ms.assetid: 8cfd1fc7-b6a8-5d1a-9dc8-58ca5521d3a8
ms.date: 06/08/2017
---


# ChartObjects.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **ChartObjects** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[ChartObjects Object](chartobjects-object-excel.md)

