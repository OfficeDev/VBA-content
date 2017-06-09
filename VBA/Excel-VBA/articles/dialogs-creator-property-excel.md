---
title: Dialogs.Creator Property (Excel)
keywords: vbaxl10.chm253074
f1_keywords:
- vbaxl10.chm253074
ms.prod: excel
api_name:
- Excel.Dialogs.Creator
ms.assetid: 4685d784-ba3f-6543-1e5e-dba7b6d6a088
ms.date: 06/08/2017
---


# Dialogs.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **Dialogs** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[Dialogs Object](dialogs-object-excel.md)

