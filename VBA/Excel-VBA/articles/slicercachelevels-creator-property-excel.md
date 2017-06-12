---
title: SlicerCacheLevels.Creator Property (Excel)
keywords: vbaxl10.chm898074
f1_keywords:
- vbaxl10.chm898074
ms.prod: excel
api_name:
- Excel.SlicerCacheLevels.Creator
ms.assetid: dfbed228-a769-86b4-7f1f-fbe55060fead
ms.date: 06/08/2017
---


# SlicerCacheLevels.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **[SlicerCacheLevels](slicercachelevels-object-excel.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[SlicerCacheLevels Object](slicercachelevels-object-excel.md)

