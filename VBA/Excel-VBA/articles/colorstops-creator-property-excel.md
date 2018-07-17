---
title: ColorStops.Creator Property (Excel)
keywords: vbaxl10.chm852074
f1_keywords:
- vbaxl10.chm852074
ms.prod: excel
api_name:
- Excel.ColorStops.Creator
ms.assetid: 9eb3106a-fb64-ba9a-8bf1-fc7ed2a3eb0e
ms.date: 06/08/2017
---


# ColorStops.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**


## Syntax

 _expression_ . **Creator**

 _expression_ An expression that returns a **ColorStops** object.


### Return Value

XlCreator


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The Creator property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL. 


## See also


#### Concepts


[ColorStops Object](colorstops-object-excel.md)

