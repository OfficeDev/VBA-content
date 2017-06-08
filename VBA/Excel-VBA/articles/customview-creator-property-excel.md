---
title: CustomView.Creator Property (Excel)
keywords: vbaxl10.chm507074
f1_keywords:
- vbaxl10.chm507074
ms.prod: excel
api_name:
- Excel.CustomView.Creator
ms.assetid: cc98f447-aa59-e2ed-c516-558d877ef791
ms.date: 06/08/2017
---


# CustomView.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **CustomView** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[CustomView Object](customview-object-excel.md)

