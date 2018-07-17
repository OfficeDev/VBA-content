---
title: OLEObjects.Creator Property (Excel)
keywords: vbaxl10.chm418074
f1_keywords:
- vbaxl10.chm418074
ms.prod: excel
api_name:
- Excel.OLEObjects.Creator
ms.assetid: b84107a4-d94c-a2b1-0a70-c4515b1d1da2
ms.date: 06/08/2017
---


# OLEObjects.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents an **OLEObjects** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[OLEObjects Object](oleobjects-object-excel.md)

