---
title: DropLines.Creator Property (Excel)
keywords: vbaxl10.chm603074
f1_keywords:
- vbaxl10.chm603074
ms.prod: excel
api_name:
- Excel.DropLines.Creator
ms.assetid: c1c7acab-e33b-ec58-6303-c31923c3f1fc
ms.date: 06/08/2017
---


# DropLines.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **DropLines** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[DropLines Object](droplines-object-excel.md)

