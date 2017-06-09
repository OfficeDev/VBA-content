---
title: AutoRecover.Creator Property (Excel)
keywords: vbaxl10.chm695074
f1_keywords:
- vbaxl10.chm695074
ms.prod: excel
api_name:
- Excel.AutoRecover.Creator
ms.assetid: 4c0849f0-e27d-de8f-0916-12ef450b10c9
ms.date: 06/08/2017
---


# AutoRecover.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents an **AutoRecover** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[AutoRecover Object](autorecover-object-excel.md)

