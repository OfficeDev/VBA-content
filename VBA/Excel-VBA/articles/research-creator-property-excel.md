---
title: Research.Creator Property (Excel)
keywords: vbaxl10.chm848074
f1_keywords:
- vbaxl10.chm848074
ms.prod: excel
api_name:
- Excel.Research.Creator
ms.assetid: b2fb9ca3-00a0-036b-7f9d-ac16a1367637
ms.date: 06/08/2017
---


# Research.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **Research** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[Research Object](research-object-excel.md)

