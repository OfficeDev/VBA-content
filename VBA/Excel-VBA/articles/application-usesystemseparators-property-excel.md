---
title: Application.UseSystemSeparators Property (Excel)
keywords: vbaxl10.chm133290
f1_keywords:
- vbaxl10.chm133290
ms.prod: excel
api_name:
- Excel.Application.UseSystemSeparators
ms.assetid: eefa7bd0-9633-2f8a-cc80-61b1649fbace
ms.date: 06/08/2017
---


# Application.UseSystemSeparators Property (Excel)

 **True** (default) if the system separators of Microsoft Excel are enabled. Read/write **Boolean** .


## Syntax

 _expression_ . **UseSystemSeparators**

 _expression_ A variable that represents an **Application** object.


## Example

In this example, "1,234,567.89" is placed in cell A1. The system separators are then changed to dashes for the decimals and thousands separators.


```vb
Sub ChangeSystemSeparators() 
 
 Range("A1").Formula = "1,234,567.89" 
 MsgBox "The system separators will now change." 
 
 ' Define separators and apply. 
 Application.DecimalSeparator = "-" 
 Application.ThousandsSeparator = "-" 
 Application.UseSystemSeparators = False 
 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

