---
title: GroupShapes.Creator Property (Excel)
keywords: vbaxl10.chm641074
f1_keywords:
- vbaxl10.chm641074
ms.prod: excel
api_name:
- Excel.GroupShapes.Creator
ms.assetid: b93ba7af-2a53-591e-18b5-3eb645a96f7a
ms.date: 06/08/2017
---


# GroupShapes.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **GroupShapes** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[GroupShapes Object](groupshapes-object-excel.md)

