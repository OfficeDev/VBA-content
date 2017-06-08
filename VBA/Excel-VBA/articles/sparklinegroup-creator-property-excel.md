---
title: SparklineGroup.Creator Property (Excel)
keywords: vbaxl10.chm870074
f1_keywords:
- vbaxl10.chm870074
ms.prod: excel
api_name:
- Excel.SparklineGroup.Creator
ms.assetid: 8a6a55f2-169f-4c65-e52c-9c182421cf4d
ms.date: 06/08/2017
---


# SparklineGroup.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **[SparklineGroup](sparklinegroup-object-excel.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[SparklineGroup Object](sparklinegroup-object-excel.md)

