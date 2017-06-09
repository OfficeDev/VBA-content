---
title: PivotTable.TableRange2 Property (Excel)
keywords: vbaxl10.chm235099
f1_keywords:
- vbaxl10.chm235099
ms.prod: excel
api_name:
- Excel.PivotTable.TableRange2
ms.assetid: 7a1ab832-baa1-f461-7036-53a0593695e7
ms.date: 06/08/2017
---


# PivotTable.TableRange2 Property (Excel)

Returns a  **[Range](range-object-excel.md)** object that represents the range containing the entire PivotTable report, including page fields. Read-only.


## Syntax

 _expression_ . **TableRange2**

 _expression_ A variable that represents a **PivotTable** object.


## Remarks

The  **[TableRange1](pivottable-tablerange1-property-excel.md)** property doesn't include page fields.


## Example

This example selects the entire PivotTable report, including its page fields.


```vb
Worksheets("Sheet1").Activate 
Range("A3").PivotTable.TableRange2.Select 

```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

