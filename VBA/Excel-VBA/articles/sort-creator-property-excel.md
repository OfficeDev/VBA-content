---
title: Sort.Creator Property (Excel)
keywords: vbaxl10.chm846074
f1_keywords:
- vbaxl10.chm846074
ms.prod: excel
api_name:
- Excel.Sort.Creator
ms.assetid: 578f0917-6778-e3df-7935-2c1121536f60
ms.date: 06/08/2017
---


# Sort.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **Sort** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[Sort Object](sort-object-excel.md)

