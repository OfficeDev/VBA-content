---
title: CustomProperties.Creator Property (Excel)
keywords: vbaxl10.chm679074
f1_keywords:
- vbaxl10.chm679074
ms.prod: excel
api_name:
- Excel.CustomProperties.Creator
ms.assetid: f40d5ca1-0606-e3ec-e4b3-278ec4f0e5f6
ms.date: 06/08/2017
---


# CustomProperties.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **CustomProperties** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[CustomProperties Object](customproperties-object-excel.md)

