---
title: MultiThreadedCalculation.Creator Property (Excel)
keywords: vbaxl10.chm858074
f1_keywords:
- vbaxl10.chm858074
ms.prod: excel
api_name:
- Excel.MultiThreadedCalculation.Creator
ms.assetid: 4121064b-2a70-e46c-c4e0-dc72cb894edf
ms.date: 06/08/2017
---


# MultiThreadedCalculation.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **MultiThreadedCalculation** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[MultiThreadedCalculation Object](multithreadedcalculation-object-excel.md)

