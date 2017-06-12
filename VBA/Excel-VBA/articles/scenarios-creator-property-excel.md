---
title: Scenarios.Creator Property (Excel)
keywords: vbaxl10.chm361074
f1_keywords:
- vbaxl10.chm361074
ms.prod: excel
api_name:
- Excel.Scenarios.Creator
ms.assetid: 90485e33-0812-a678-1ec1-865271450264
ms.date: 06/08/2017
---


# Scenarios.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **Scenarios** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[Scenarios Object](scenarios-object-excel.md)

