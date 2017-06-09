---
title: Application.ReferenceStyle Property (Excel)
keywords: vbaxl10.chm133197
f1_keywords:
- vbaxl10.chm133197
ms.prod: excel
api_name:
- Excel.Application.ReferenceStyle
ms.assetid: 86c4931b-ab1a-0363-d048-5195707a952b
ms.date: 06/08/2017
---


# Application.ReferenceStyle Property (Excel)

Returns or sets how Microsoft Excel displays cell references and row and column headings in either A1 or R1C1 reference style. Read/write  **[XlReferenceStyle](xlreferencestyle-enumeration-excel.md)** .


## Syntax

 _expression_ . **ReferenceStyle**

 _expression_ A variable that represents an **Application** object.


## Remarks





| **XlReferenceStyle** can be one of these **XlReferenceStyle** constants.|
| **xlA1**|
| **xlR1C1**|

## Example

This example displays the current reference style.


```vb
If Application.ReferenceStyle = xlR1C1 Then 
 MsgBox ("Microsoft Excel is using R1C1 references") 
Else 
 MsgBox ("Microsoft Excel is using A1 references") 
End If
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

