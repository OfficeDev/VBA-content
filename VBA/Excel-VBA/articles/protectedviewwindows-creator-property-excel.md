---
title: ProtectedViewWindows.Creator Property (Excel)
keywords: vbaxl10.chm912074
f1_keywords:
- vbaxl10.chm912074
ms.prod: excel
api_name:
- Excel.ProtectedViewWindows.Creator
ms.assetid: f1c6f32e-57dc-3a3c-0d6f-f43f94c0f39f
ms.date: 06/08/2017
---


# ProtectedViewWindows.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **ProtectedViewWindows** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[ProtectedViewWindows Object](protectedviewwindows-object-excel.md)

