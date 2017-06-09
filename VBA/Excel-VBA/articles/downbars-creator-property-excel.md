---
title: DownBars.Creator Property (Excel)
keywords: vbaxl10.chm609074
f1_keywords:
- vbaxl10.chm609074
ms.prod: excel
api_name:
- Excel.DownBars.Creator
ms.assetid: 157413a9-f3f7-8d98-294c-8580dfabdd2b
ms.date: 06/08/2017
---


# DownBars.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **DownBars** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[DownBars Object](downbars-object-excel.md)

