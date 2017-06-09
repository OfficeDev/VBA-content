---
title: Parameters.Creator Property (Excel)
keywords: vbaxl10.chm524074
f1_keywords:
- vbaxl10.chm524074
ms.prod: excel
api_name:
- Excel.Parameters.Creator
ms.assetid: 357ca5be-2f41-4bac-a10a-b917441f6e29
ms.date: 06/08/2017
---


# Parameters.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **Parameters** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[Parameters Object](parameters-object-excel.md)

