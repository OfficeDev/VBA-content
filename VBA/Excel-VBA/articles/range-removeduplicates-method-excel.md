---
title: Range.RemoveDuplicates Method (Excel)
keywords: vbaxl10.chm144243
f1_keywords:
- vbaxl10.chm144243
ms.prod: excel
api_name:
- Excel.Range.RemoveDuplicates
ms.assetid: 0e74bde2-08b3-898d-0b30-53de911bd7e9
ms.date: 06/08/2017
---


# Range.RemoveDuplicates Method (Excel)

Removes duplicate values from a range of values.


## Syntax

 _expression_ . **RemoveDuplicates**( **_Columns_** , **_Header_** )

 _expression_ A variable that represents a **Range** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Columns_|Required| **Variant**|Array of indexes of the columns that contain the duplicate information. |
| _Header_|Optional| **XlYesNoGuess**|Specifies whether the first row contains header information.  **xlNo** is the default value; specify **xlGuess** if you want Excel to attempt to determine the header.|

## Example

The following code sample removes duplicates with the first 2 columns


```vb
ActiveSheet.Range("A1:C100").RemoveDuplicates Columns:=Array(1,2), Header:=xlYes
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

