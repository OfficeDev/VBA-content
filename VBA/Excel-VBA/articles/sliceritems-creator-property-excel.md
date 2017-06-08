---
title: SlicerItems.Creator Property (Excel)
keywords: vbaxl10.chm908074
f1_keywords:
- vbaxl10.chm908074
ms.prod: excel
api_name:
- Excel.SlicerItems.Creator
ms.assetid: d7002e14-3c07-3255-6b01-556fc1d3c503
ms.date: 06/08/2017
---


# SlicerItems.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only.


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **[SlicerItems](sliceritems-object-excel.md)** object.


### Return Value

 **Long**


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[SlicerItems Object](sliceritems-object-excel.md)

