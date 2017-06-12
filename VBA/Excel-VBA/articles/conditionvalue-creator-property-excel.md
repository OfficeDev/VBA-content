---
title: ConditionValue.Creator Property (Excel)
keywords: vbaxl10.chm803074
f1_keywords:
- vbaxl10.chm803074
ms.prod: excel
api_name:
- Excel.ConditionValue.Creator
ms.assetid: 74c0263a-5f2a-3a44-b3ff-4a5b7cddf13a
ms.date: 06/08/2017
---


# ConditionValue.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **ConditionValue** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[ConditionValue Object](conditionvalue-object-excel.md)

