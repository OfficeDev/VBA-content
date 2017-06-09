---
title: DisplayFormat.Creator Property (Excel)
keywords: vbaxl10.chm892074
f1_keywords:
- vbaxl10.chm892074
ms.prod: excel
api_name:
- Excel.DisplayFormat.Creator
ms.assetid: 6e3749be-adec-bb6c-dc24-232e5046ef12
ms.date: 06/08/2017
---


# DisplayFormat.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **[DisplayFormat](displayformat-object-excel.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[DisplayFormat Object](displayformat-object-excel.md)

