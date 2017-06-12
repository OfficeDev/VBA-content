---
title: CalculatedMember.Creator Property (Excel)
keywords: vbaxl10.chm685074
f1_keywords:
- vbaxl10.chm685074
ms.prod: excel
api_name:
- Excel.CalculatedMember.Creator
ms.assetid: 2892e70d-6c8d-b327-138c-80fa0222a375
ms.date: 06/08/2017
---


# CalculatedMember.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **CalculatedMember** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[CalculatedMember Object](calculatedmember-object-excel.md)

