---
title: SparkColor.Creator Property (Excel)
keywords: vbaxl10.chm882074
f1_keywords:
- vbaxl10.chm882074
ms.prod: excel
api_name:
- Excel.SparkColor.Creator
ms.assetid: 4acfe022-4841-70b1-c38b-dd535e9cba9b
ms.date: 06/08/2017
---


# SparkColor.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **[SparkColor](sparkcolor-object-excel.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[SparkColor Object](sparkcolor-object-excel.md)

