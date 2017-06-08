---
title: Parameter.Creator Property (Excel)
keywords: vbaxl10.chm522074
f1_keywords:
- vbaxl10.chm522074
ms.prod: excel
api_name:
- Excel.Parameter.Creator
ms.assetid: 3af59d13-b371-3e9f-b6d2-62452a2cba98
ms.date: 06/08/2017
---


# Parameter.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **Parameter** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[Parameter Object](parameter-object-excel.md)

