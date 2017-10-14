---
title: AxisTitle.Creator Property (Excel)
keywords: vbaxl10.chm564074
f1_keywords:
- vbaxl10.chm564074
ms.prod: excel
api_name:
- Excel.AxisTitle.Creator
ms.assetid: 1a1ba9e2-f3fb-d1d1-965e-b236da4564b4
ms.date: 06/08/2017
---


# AxisTitle.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents an **AxisTitle** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[AxisTitle Object](axistitle-object-excel.md)

