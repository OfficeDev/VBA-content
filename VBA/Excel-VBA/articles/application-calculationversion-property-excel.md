---
title: Application.CalculationVersion Property (Excel)
keywords: vbaxl10.chm133257
f1_keywords:
- vbaxl10.chm133257
ms.prod: excel
api_name:
- Excel.Application.CalculationVersion
ms.assetid: 10de3816-9873-09e5-4141-effdbfe5cd9c
ms.date: 06/08/2017
---


# Application.CalculationVersion Property (Excel)

Returns a number whose rightmost four digits are the minor calculation engine version number, and whose other digits (on the left) are the major version of Microsoft Excel. Read-only  **Long** .


## Syntax

 _expression_ . **CalculationVersion**

 _expression_ A variable that represents an **Application** object.


## Remarks

If the workbook was saved in an earlier version of Excel and if the workbook hasn't been fully recalculated, then this property returns 0.


## Example

This example compares the version of Microsoft Excel with the version of Excel that the workbook was last calculated in. If the two version numbers are different, the example sets the  `blnFullCalc` variable to **True** .


```vb
If Application.CalculationVersion <> _ 
 Workbooks(1).CalculationVersion Then 
 blnFullCalc = True 
Else 
 blnFullCalc = False 
End If
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

