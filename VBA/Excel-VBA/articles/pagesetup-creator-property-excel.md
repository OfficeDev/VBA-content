---
title: PageSetup.Creator Property (Excel)
keywords: vbaxl10.chm472074
f1_keywords:
- vbaxl10.chm472074
ms.prod: excel
api_name:
- Excel.PageSetup.Creator
ms.assetid: 88f7b9ab-0176-9495-9d1a-57b8a78e5e3b
ms.date: 06/08/2017
---


# PageSetup.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **PageSetup** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[PageSetup Object](pagesetup-object-excel.md)

