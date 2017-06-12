---
title: CalculatedMembers.Creator Property (Excel)
keywords: vbaxl10.chm683074
f1_keywords:
- vbaxl10.chm683074
ms.prod: excel
api_name:
- Excel.CalculatedMembers.Creator
ms.assetid: a743dd00-e7f5-236f-989c-e1e442389d08
ms.date: 06/08/2017
---


# CalculatedMembers.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **CalculatedMembers** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[CalculatedMembers Collection](calculatedmembers-object-excel.md)

