---
title: Range.AddressLocal Property (Excel)
keywords: vbaxl10.chm144077
f1_keywords:
- vbaxl10.chm144077
ms.prod: excel
api_name:
- Excel.Range.AddressLocal
ms.assetid: 20332d15-dc37-1819-472f-ef208d8b3a5b
ms.date: 06/08/2017
---


# Range.AddressLocal Property (Excel)

Returns the range reference for the specified range in the language of the user. Read-only  **String** .


## Syntax

 _expression_ . **AddressLocal**( **_RowAbsolute_** , **_ColumnAbsolute_** , **_ReferenceStyle_** , **_External_** , **_RelativeTo_** )

 _expression_ A variable that represents a **Range** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _RowAbsolute_|Optional| **Variant**| **True** to return the row part of the reference as an absolute reference. The default value is **True** .|
| _ColumnAbsolute_|Optional| **Variant**| **True** to return the column part of the reference as an absolute reference. The default value is **True** .|
| _ReferenceStyle_|Optional| **[XlReferenceStyle](xlreferencestyle-enumeration-excel.md)**|One of the constants for  **XlReferenceStyle** specifying A1-style or R1C1-style reference.|
| _External_|Optional| **Variant**| **True** to return an external reference. **False** to return a local reference. The default value is **False** .|
| _RelativeTo_|Optional| **Variant**|If  _RowAbsolute_ and _ColumnAbsolute_ are both set to **False** and _ReferenceStyle_ is set to **xlR1C1** , you must include a starting point for the relative reference. This argument is a **Range** object that defines the starting point for the reference.|

## Remarks

If the reference contains more than one cell,  _RowAbsolute_ and _ColumnAbsolute_ apply to all rows and all columns, respectively.


## Example

Assume that this example was created using U.S. English language support and was then run in using German language support. The example displays the text shown in the comments.


```vb
Set mc = Worksheets(1).Cells(1, 1) 
MsgBox mc.AddressLocal() ' $A$1 
MsgBox mc.AddressLocal(RowAbsolute:=False) ' $A1 
MsgBox mc.AddressLocal(ReferenceStyle:=xlR1C1) ' Z1S1 
MsgBox mc.AddressLocal(ReferenceStyle:=xlR1C1, _ 
 RowAbsolute:=False, _ 
 ColumnAbsolute:=False, _ 
 RelativeTo:=Worksheets(1).Cells(3, 3)) ' Z(-2)S(-2)
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

