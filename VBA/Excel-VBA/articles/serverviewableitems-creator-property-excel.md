---
title: ServerViewableItems.Creator Property (Excel)
keywords: vbaxl10.chm832074
f1_keywords:
- vbaxl10.chm832074
ms.prod: excel
api_name:
- Excel.ServerViewableItems.Creator
ms.assetid: ebc56118-1d24-45ee-b2a1-2fc59095a4e7
ms.date: 06/08/2017
---


# ServerViewableItems.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **ServerViewableItems** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[ServerViewableItems Object](serverviewableitems-object-excel.md)

