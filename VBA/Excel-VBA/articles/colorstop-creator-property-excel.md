---
title: ColorStop.Creator Property (Excel)
keywords: vbaxl10.chm850074
f1_keywords:
- vbaxl10.chm850074
ms.prod: excel
api_name:
- Excel.ColorStop.Creator
ms.assetid: 99789f97-d576-1be6-40c5-9cd2a5984751
ms.date: 06/08/2017
---


# ColorStop.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **ColorStop** object.


### Return Value

XlCreator


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The Creator property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL. 


## See also


#### Concepts


[ColorStop Object](colorstop-object-excel.md)

