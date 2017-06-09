---
title: Application.CalculateFullRebuild Method (Excel)
keywords: vbaxl10.chm133272
f1_keywords:
- vbaxl10.chm133272
ms.prod: excel
api_name:
- Excel.Application.CalculateFullRebuild
ms.assetid: 6d3dac24-7fb8-05fd-b6ee-cb3ef7d5f33a
ms.date: 06/08/2017
---


# Application.CalculateFullRebuild Method (Excel)

For all open workbooks, forces a full calculation of the data and rebuilds the dependencies.


## Syntax

 _expression_ . **CalculateFullRebuild**

 _expression_ A variable that represents an **Application** object.


## Remarks

Dependencies are the formulas that depend on other cells. For example, the formula "=A1" depends on cell A1. The  **CalculateFullRebuild** method is similar to re-entering all formulas.


## Example

This example compares the version of Microsoft Excel with the version of Excel in which the workbook was last calculated. If the two version numbers are different, a full calculation of the data in all open workbooks is performed and the dependencies are rebuilt.


```vb
Sub UseCalculateFullRebuild() 
 
 If Application.CalculationVersion <> _ 
 Workbooks(1).CalculationVersion Then 
 Application.CalculateFullRebuild 
 End If 
 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

