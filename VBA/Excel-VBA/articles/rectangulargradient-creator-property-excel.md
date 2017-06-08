---
title: RectangularGradient.Creator Property (Excel)
keywords: vbaxl10.chm856074
f1_keywords:
- vbaxl10.chm856074
ms.prod: excel
api_name:
- Excel.RectangularGradient.Creator
ms.assetid: b6697d8f-7cd0-a731-3375-e18901f55ed0
ms.date: 06/08/2017
---


# RectangularGradient.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **RectangularGradient** object.


### Return Value

XlCreator


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The Creator property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL. 


## See also


#### Concepts


[RectangularGradient Object](rectangulargradient-object-excel.md)

