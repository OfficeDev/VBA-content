---
title: Application.ConstrainNumeric Property (Excel)
keywords: vbaxl10.chm133096
f1_keywords:
- vbaxl10.chm133096
ms.prod: excel
api_name:
- Excel.Application.ConstrainNumeric
ms.assetid: 910dd5ad-1750-71b8-8c12-df5107d21063
ms.date: 06/08/2017
---


# Application.ConstrainNumeric Property (Excel)

 **True** if handwriting recognition is limited to numbers and punctuation only. Read/write **Boolean** .


## Syntax

 _expression_ . **ConstrainNumeric**

 _expression_ A variable that represents an **Application** object.


## Remarks

This property is available only if you're using Microsoft Windows for Pen Computing. If you try to set this property under any other operating system, an error occurs.


## Example

This example limits handwriting recognition to numbers and punctuation only if Microsoft Windows for Pen Computing is running.


```vb
If Application.WindowsForPens Then 
 Application.ConstrainNumeric = True 
End If
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

