---
title: CustomViews.Creator Property (Excel)
keywords: vbaxl10.chm505074
f1_keywords:
- vbaxl10.chm505074
ms.prod: excel
api_name:
- Excel.CustomViews.Creator
ms.assetid: c0d96d50-e126-09cc-3660-e2f0dc1fb566
ms.date: 06/08/2017
---


# CustomViews.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **CustomViews** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[CustomViews Object](customviews-object-excel.md)

