---
title: Points.Creator Property (Excel)
keywords: vbaxl10.chm573074
f1_keywords:
- vbaxl10.chm573074
ms.prod: excel
api_name:
- Excel.Points.Creator
ms.assetid: 2924d441-34b8-6a19-9591-57a2824248d5
ms.date: 06/08/2017
---


# Points.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **Points** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[Points Object](points-object-excel.md)

