---
title: ErrorBars.Creator Property (Excel)
keywords: vbaxl10.chm623074
f1_keywords:
- vbaxl10.chm623074
ms.prod: excel
api_name:
- Excel.ErrorBars.Creator
ms.assetid: 8a54a5dd-a62d-e027-8c44-ba4f97ac425d
ms.date: 06/08/2017
---


# ErrorBars.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents an **ErrorBars** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[ErrorBars Object](errorbars-object-excel.md)

