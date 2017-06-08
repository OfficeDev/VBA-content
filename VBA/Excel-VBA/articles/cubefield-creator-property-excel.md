---
title: CubeField.Creator Property (Excel)
keywords: vbaxl10.chm667074
f1_keywords:
- vbaxl10.chm667074
ms.prod: excel
api_name:
- Excel.CubeField.Creator
ms.assetid: 2534f870-90cd-e3ab-b1fd-d63455a75809
ms.date: 06/08/2017
---


# CubeField.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **CubeField** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[CubeField Object](cubefield-object-excel.md)

