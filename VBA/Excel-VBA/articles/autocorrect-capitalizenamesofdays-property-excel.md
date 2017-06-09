---
title: AutoCorrect.CapitalizeNamesOfDays Property (Excel)
keywords: vbaxl10.chm545074
f1_keywords:
- vbaxl10.chm545074
ms.prod: excel
api_name:
- Excel.AutoCorrect.CapitalizeNamesOfDays
ms.assetid: 218f9820-8cb1-438d-7c81-4a9c4385a275
ms.date: 06/08/2017
---


# AutoCorrect.CapitalizeNamesOfDays Property (Excel)

 **True** if the first letter of day names is capitalized automatically. Read/write **Boolean** .


## Syntax

 _expression_ . **CapitalizeNamesOfDays**

 _expression_ A variable that represents an **AutoCorrect** object.


## Example

This example sets Microsoft Excel to capitalize the first letter of the names of days.


```vb
With Application.AutoCorrect 
 .CapitalizeNamesOfDays = True 
 .ReplaceText = True 
End With
```


## See also


#### Concepts


[AutoCorrect Object](autocorrect-object-excel.md)

