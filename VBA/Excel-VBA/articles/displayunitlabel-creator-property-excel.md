---
title: DisplayUnitLabel.Creator Property (Excel)
keywords: vbaxl10.chm672074
f1_keywords:
- vbaxl10.chm672074
ms.prod: excel
api_name:
- Excel.DisplayUnitLabel.Creator
ms.assetid: 8150429a-3533-6681-36cf-22db6196610f
ms.date: 06/08/2017
---


# DisplayUnitLabel.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **DisplayUnitLabel** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[DisplayUnitLabel Object](displayunitlabel-object-excel.md)

