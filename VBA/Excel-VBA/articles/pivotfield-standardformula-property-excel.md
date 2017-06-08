---
title: PivotField.StandardFormula Property (Excel)
keywords: vbaxl10.chm240128
f1_keywords:
- vbaxl10.chm240128
ms.prod: excel
api_name:
- Excel.PivotField.StandardFormula
ms.assetid: 14d5cd3e-29d8-a70a-b52b-41c42252ef7c
ms.date: 06/08/2017
---


# PivotField.StandardFormula Property (Excel)

Returns or sets a  **String** specifying formulas with standard English (United States) formatting. Read/write.


## Syntax

 _expression_ . **StandardFormula**

 _expression_ A variable that represents a **PivotField** object.


## Remarks

The  **StandardFormula** property primarily affects item names with date or number formatting. It provides a way to specify or query a formula for a given calculated item.

The  **[StandardFormula](pivotfield-standardformula-property-excel.md)** property is "international-friendly" whereas the **[Formula](pivotfield-formula-property-excel.md)** property is not.


## Example

This example adds 10 to the Decimals field and displays it as a calculated item in the data field. The example assumes that a PivotTable exists on the active worksheet and that a field titled "Decimals" exists in the data table.


```vb
Sub UseStandardFomula() 
 
 Dim pvtTable As PivotTable 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 
 ' Change calculated field of decimals by adding '10'. 
 pvtTable.CalculatedFields.Item(1).StandardFormula = "Decimals + 10" 
 
End Sub
```


## See also


#### Concepts


[PivotField Object](pivotfield-object-excel.md)

