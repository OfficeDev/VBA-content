---
title: LeaderLines.Creator Property (Excel)
keywords: vbaxl10.chm605074
f1_keywords:
- vbaxl10.chm605074
ms.prod: excel
api_name:
- Excel.LeaderLines.Creator
ms.assetid: 6b2769a0-b54d-e75d-2ada-aa220809e238
ms.date: 06/08/2017
---


# LeaderLines.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **LeaderLines** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[LeaderLines Object](leaderlines-object-excel.md)

