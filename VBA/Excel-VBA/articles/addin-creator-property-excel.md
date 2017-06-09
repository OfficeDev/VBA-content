---
title: AddIn.Creator Property (Excel)
keywords: vbaxl10.chm184074
f1_keywords:
- vbaxl10.chm184074
ms.prod: excel
api_name:
- Excel.AddIn.Creator
ms.assetid: d3c90243-1963-15d5-9ee1-1cd780ee20a0
ms.date: 06/08/2017
---


# AddIn.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ An expression that returns a **AddIn** object.


### Return Value

XlCreator


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[AddIn Object](addin-object-excel.md)

