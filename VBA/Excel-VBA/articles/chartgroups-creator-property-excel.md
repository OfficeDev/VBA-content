---
title: ChartGroups.Creator Property (Excel)
keywords: vbaxl10.chm569074
f1_keywords:
- vbaxl10.chm569074
ms.prod: excel
api_name:
- Excel.ChartGroups.Creator
ms.assetid: 3008f9c3-a7d8-3202-2edb-d090deb039af
ms.date: 06/08/2017
---


# ChartGroups.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **ChartGroups** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[ChartGroups Object](chartgroups-object-excel.md)

