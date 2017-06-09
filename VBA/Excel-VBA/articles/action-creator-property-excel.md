---
title: Action.Creator Property (Excel)
keywords: vbaxl10.chm797074
f1_keywords:
- vbaxl10.chm797074
ms.prod: excel
api_name:
- Excel.Action.Creator
ms.assetid: 4a7459d1-64aa-07b0-c0e5-56a0d837c8d2
ms.date: 06/08/2017
---


# Action.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ An expression that returns a **Action** object.


### Return Value

XlCreator


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[Action Object](action-object-excel.md)

