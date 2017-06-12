---
title: UpBars.Creator Property (Excel)
keywords: vbaxl10.chm607074
f1_keywords:
- vbaxl10.chm607074
ms.prod: excel
api_name:
- Excel.UpBars.Creator
ms.assetid: 0cfa8ef1-fd09-073c-8156-f2c684894b90
ms.date: 06/08/2017
---


# UpBars.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **UpBars** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[UpBars Object](upbars-object-excel.md)

