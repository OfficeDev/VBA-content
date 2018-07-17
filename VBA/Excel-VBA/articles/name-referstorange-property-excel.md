---
title: Name.RefersToRange Property (Excel)
keywords: vbaxl10.chm490088
f1_keywords:
- vbaxl10.chm490088
ms.prod: excel
api_name:
- Excel.Name.RefersToRange
ms.assetid: 81c0e2fe-8ce6-0df9-9ffa-0931b87487e7
ms.date: 06/08/2017
---


# Name.RefersToRange Property (Excel)

Returns the  **[Range](range-object-excel.md)** object referred to by a **Name** object. Read-only.


## Syntax

 _expression_ . **RefersToRange**

 _expression_ A variable that represents a **Name** object.


## Remarks

If the  **Name** object doesn't refer to a range (for example, if it refers to a constant or a formula), this property fails.

To change the range that a name refers to, use the  **[RefersTo](name-refersto-property-excel.md)** property.


## Example

This example displays the number of rows and columns in the print area on the active worksheet.


 **Note**  Ensure that you establish a print area on the active sheet of the current workbook.


```
p = Sheets(ActiveSheet.Name).Names("Print_Area").RefersToRange.Value 
MsgBox "Print_Area: " &; UBound(p, 1) &; " rows, " &; _ 
 UBound(p, 2) &; " columns"
```


## See also


#### Concepts


[Name Object](name-object-excel.md)

