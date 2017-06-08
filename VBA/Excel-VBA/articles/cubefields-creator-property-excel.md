---
title: CubeFields.Creator Property (Excel)
keywords: vbaxl10.chm669074
f1_keywords:
- vbaxl10.chm669074
ms.prod: excel
api_name:
- Excel.CubeFields.Creator
ms.assetid: 11680e70-3280-7cb4-ef21-390653e5adb9
ms.date: 06/08/2017
---


# CubeFields.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **CubeFields** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[CubeFields Object](cubefields-object-excel.md)

