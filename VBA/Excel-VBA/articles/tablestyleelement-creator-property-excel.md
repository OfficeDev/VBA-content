---
title: TableStyleElement.Creator Property (Excel)
keywords: vbaxl10.chm834074
f1_keywords:
- vbaxl10.chm834074
ms.prod: excel
api_name:
- Excel.TableStyleElement.Creator
ms.assetid: ab9524d1-7d61-cc43-2d8f-0b087f1ccb1b
ms.date: 06/08/2017
---


# TableStyleElement.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **TableStyleElement** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[TableStyleElement Object](tablestyleelement-object-excel.md)

