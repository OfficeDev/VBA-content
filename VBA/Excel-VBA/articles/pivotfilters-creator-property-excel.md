---
title: PivotFilters.Creator Property (Excel)
keywords: vbaxl10.chm771074
f1_keywords:
- vbaxl10.chm771074
ms.prod: excel
api_name:
- Excel.PivotFilters.Creator
ms.assetid: f20c1952-90de-3d14-5d31-77f44ce24767
ms.date: 06/08/2017
---


# PivotFilters.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **PivotFilters** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[PivotFilters Object](pivotfilters-object-excel.md)

