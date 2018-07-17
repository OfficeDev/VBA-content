---
title: ChartArea.Creator Property (Excel)
keywords: vbaxl10.chm619074
f1_keywords:
- vbaxl10.chm619074
ms.prod: excel
api_name:
- Excel.ChartArea.Creator
ms.assetid: 430863d6-d88f-06a3-f979-6f48d2c551f4
ms.date: 06/08/2017
---


# ChartArea.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **ChartArea** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[ChartArea Object](chartarea-object-excel.md)

