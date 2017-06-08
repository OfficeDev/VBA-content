---
title: RecentFiles.Creator Property (Excel)
keywords: vbaxl10.chm171074
f1_keywords:
- vbaxl10.chm171074
ms.prod: excel
api_name:
- Excel.RecentFiles.Creator
ms.assetid: 83b6210e-5994-2468-f4b9-0884abc689fc
ms.date: 06/08/2017
---


# RecentFiles.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **RecentFiles** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[RecentFiles Object](recentfiles-object-excel.md)

