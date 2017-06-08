---
title: WebOptions.Creator Property (Excel)
keywords: vbaxl10.chm661074
f1_keywords:
- vbaxl10.chm661074
ms.prod: excel
api_name:
- Excel.WebOptions.Creator
ms.assetid: 506df7ba-2e4f-69af-793c-96c5f2aa2f1c
ms.date: 06/08/2017
---


# WebOptions.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **WebOptions** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[WebOptions Object](weboptions-object-excel.md)

