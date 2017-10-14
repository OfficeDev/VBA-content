---
title: ChartGroup.Creator Property (Excel)
keywords: vbaxl10.chm567074
f1_keywords:
- vbaxl10.chm567074
ms.prod: excel
api_name:
- Excel.ChartGroup.Creator
ms.assetid: 5f1ce433-8248-47d6-ea1b-90c7c8aac75e
ms.date: 06/08/2017
---


# ChartGroup.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **ChartGroup** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[ChartGroup Object](chartgroup-object-excel.md)

