---
title: HPageBreak.Creator Property (Excel)
keywords: vbaxl10.chm158074
f1_keywords:
- vbaxl10.chm158074
ms.prod: excel
api_name:
- Excel.HPageBreak.Creator
ms.assetid: 4cefce48-34ee-40a4-0a0c-d3dade1cb065
ms.date: 06/08/2017
---


# HPageBreak.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **HPageBreak** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[HPageBreak Object](hpagebreak-object-excel.md)

