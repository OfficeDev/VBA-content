---
title: Phonetics.Creator Property (Excel)
keywords: vbaxl10.chm657074
f1_keywords:
- vbaxl10.chm657074
ms.prod: excel
api_name:
- Excel.Phonetics.Creator
ms.assetid: 7419d5c6-88f4-f07b-083a-ea15fdda3765
ms.date: 06/08/2017
---


# Phonetics.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **Phonetics** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[Phonetics Object](phonetics-object-excel.md)

