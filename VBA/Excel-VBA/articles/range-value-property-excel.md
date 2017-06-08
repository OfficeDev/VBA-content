---
title: Range.Value Property (Excel)
keywords: vbaxl10.chm144216
f1_keywords:
- vbaxl10.chm144216
ms.prod: excel
api_name:
- Excel.Range.Value
ms.assetid: 23f28b24-430a-6ea4-4895-0afff8dff218
ms.date: 06/08/2017
---


# Range.Value Property (Excel)

Returns or sets a  **Variant** value that represents the value of the specified range.


## Syntax

 _expression_ . **Value**( **_RangeValueDataType_** )

 _expression_ A variable that represents a **Range** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _RangeValueDataType_|Optional| **Variant**|The range value data type. Can be a  **[xlRangeValueDataType](xlrangevaluedatatype-enumeration-excel.md)** constant.|

## Remarks

When setting a range of cells with the contents of an XML spreadsheet file, only values of the first sheet in the workbook are used. You cannot set or get a discontiguous range of cells in the XML spreadsheet format.


## Example

This example sets the value of cell A1 on Sheet1 to 3.14159.


```vb
Worksheets("Sheet1").Range("A1").Value = 3.14159
```

This example loops on cells A1:D10 on Sheet1. If one of the cells has a value less than 0.001, the code replaces the value with 0 (zero).




```vb
For Each c in Worksheets("Sheet1").Range("A1:D10") 
 If c.Value < .001 Then 
 c.Value = 0 
 End If 
Next c
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

