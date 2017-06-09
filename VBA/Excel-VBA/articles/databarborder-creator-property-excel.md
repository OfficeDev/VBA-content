---
title: DataBarBorder.Creator Property (Excel)
keywords: vbaxl10.chm884074
f1_keywords:
- vbaxl10.chm884074
ms.prod: excel
api_name:
- Excel.DataBarBorder.Creator
ms.assetid: 2d240406-f29b-6014-4cc0-06085c9573d8
ms.date: 06/08/2017
---


# DataBarBorder.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **[DataBarBorder](databarborder-object-excel.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[DataBarBorder Object](databarborder-object-excel.md)

