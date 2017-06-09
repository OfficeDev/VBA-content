---
title: ListColumn.TotalsCalculation Property (Excel)
keywords: vbaxl10.chm738079
f1_keywords:
- vbaxl10.chm738079
ms.prod: excel
api_name:
- Excel.ListColumn.TotalsCalculation
ms.assetid: bb8c1dd1-1ee6-3ef8-8af4-2b3f24eb642d
ms.date: 06/08/2017
---


# ListColumn.TotalsCalculation Property (Excel)

Determines the type of calculation in the Totals row of the list column based on the value of the  **[XlTotalsCalculation](xltotalscalculation-enumeration-excel.md)** enumeration. Read/write.


## Syntax

 _expression_ . **TotalsCalculation**

 _expression_ A variable that represents a **ListColumn** object.


## Remarks



| **XlTotalsCalculation** can be one of these **XlTotalsCalculation** constants.|
| **xlTotalsCalculationNone**|
| **xlTotalsCalculationSum**|
| **xlTotalsCalculationAverage**|
| **xlTotalsCalculationCount**|
| **xlTotalsCalculationCountNums**|
| **xlTotalsCalculationMin**|
| **xlTotalsCalculationStdDev**|
| **xlTotalsCalculationVar**|
| **xlTotalsCalculationMax**|
The Totals row doesn't need to be showing in order to set this property. There is no fixed "default" value for this property. Excel may change the state of this property, as other columns are added or deleted.


## Example


```vb
ActiveSheet.ListColumns(1).TotalsCalculation=xlTotalsCalculationSum
```


## See also


#### Concepts


[ListColumn Object](listcolumn-object-excel.md)

