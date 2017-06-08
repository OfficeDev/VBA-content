---
title: Top10.Creator Property (Excel)
keywords: vbaxl10.chm821074
f1_keywords:
- vbaxl10.chm821074
ms.prod: excel
api_name:
- Excel.Top10.Creator
ms.assetid: 47d808f6-27f5-c8d9-97ab-1d135d25e4f7
ms.date: 06/08/2017
---


# Top10.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **Top10** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[Top10 Object](top10-object-excel.md)

