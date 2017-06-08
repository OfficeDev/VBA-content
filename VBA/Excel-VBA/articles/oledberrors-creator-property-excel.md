---
title: OLEDBErrors.Creator Property (Excel)
keywords: vbaxl10.chm655074
f1_keywords:
- vbaxl10.chm655074
ms.prod: excel
api_name:
- Excel.OLEDBErrors.Creator
ms.assetid: c2143d28-5e66-5207-7d8d-82333d9de724
ms.date: 06/08/2017
---


# OLEDBErrors.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents an **OLEDBErrors** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[OLEDBErrors Object](oledberrors-object-excel.md)

