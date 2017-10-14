---
title: Range.Address Property (Excel)
keywords: vbaxl10.chm144076
f1_keywords:
- vbaxl10.chm144076
ms.prod: excel
api_name:
- Excel.Range.Address
ms.assetid: aaa2432e-9bb1-4a48-3868-86455bc53938
ms.date: 06/08/2017
---


# Range.Address Property (Excel)

Returns a  **String** value that represents the range reference in the language of the macro.


## Syntax

 _expression_ . **Address**( **_RowAbsolute_** , **_ColumnAbsolute_** , **_ReferenceStyle_** , **_External_** , **_RelativeTo_** )

 _expression_ A variable that represents a **Range** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _RowAbsolute_|Optional| **Variant**| **True** to return the row part of the reference as an absolute reference. The default value is **True** .|
| _ColumnAbsolute_|Optional| **Variant**| **True** to return the column part of the reference as an absolute reference. The default value is **True** .|
| _ReferenceStyle_|Optional| **[XlReferenceStyle](xlreferencestyle-enumeration-excel.md)**|The reference style. The default value is  **xlA1** .|
| _External_|Optional| **Variant**| **True** to return an external reference. **False** to return a local reference. The default value is **False** .|
| _RelativeTo_|Optional| **Variant**|If  _RowAbsolute_ and _ColumnAbsolute_ are **False** , and _ReferenceStyle_ is **xlR1C1** , you must include a starting point for the relative reference. This argument is a **[Range](range-object-excel.md)** object that defines the starting point.|

## Remarks

If the reference contains more than one cell,  _RowAbsolute_ and _ColumnAbsolute_ apply to all rows and columns.




## Example

The following example displays four different representations of the same cell address on Sheet1. The comments in the example are the addresses that will be displayed in the message boxes.


```vb
Set mc = Worksheets("Sheet1").Cells(1, 1) 
MsgBox mc.Address() ' $A$1 
MsgBox mc.Address(RowAbsolute:=False) ' $A1 
MsgBox mc.Address(ReferenceStyle:=xlR1C1) ' R1C1 
MsgBox mc.Address(ReferenceStyle:=xlR1C1, _ 
 RowAbsolute:=False, _ 
 ColumnAbsolute:=False, _ 
 RelativeTo:=Worksheets(1).Cells(3, 3)) ' R[-2]C[-2]
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

