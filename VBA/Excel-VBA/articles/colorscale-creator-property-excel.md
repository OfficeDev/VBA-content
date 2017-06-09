---
title: ColorScale.Creator Property (Excel)
keywords: vbaxl10.chm805074
f1_keywords:
- vbaxl10.chm805074
ms.prod: excel
api_name:
- Excel.ColorScale.Creator
ms.assetid: 60928601-77fe-2a4b-ecd5-9b8e3adeea6a
ms.date: 06/08/2017
---


# ColorScale.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **ColorScale** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[ColorScale Object](colorscale-object-excel.md)

