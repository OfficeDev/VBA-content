---
title: Floor.Creator Property (Excel)
keywords: vbaxl10.chm611074
f1_keywords:
- vbaxl10.chm611074
ms.prod: excel
api_name:
- Excel.Floor.Creator
ms.assetid: 04cfbb36-51f5-a1d1-0f22-a1ecf9be682e
ms.date: 06/08/2017
---


# Floor.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **Floor** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[Floor Object](floor-object-excel.md)

