---
title: Range.Parse Method (Excel)
keywords: vbaxl10.chm144173
f1_keywords:
- vbaxl10.chm144173
ms.prod: excel
api_name:
- Excel.Range.Parse
ms.assetid: 3580aeb7-e868-894a-9dd5-8e37475fb267
ms.date: 06/08/2017
---


# Range.Parse Method (Excel)

Parses a range of data and breaks it into multiple cells. Distributes the contents of the range to fill several adjacent columns; the range can be no more than one column wide.


## Syntax

 _expression_ . **Parse**( **_ParseLine_** , **_Destination_** )

 _expression_ A variable that represents a **Range** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ParseLine_|Optional| **Variant**|A string that contains left and right brackets to indicate where the cells should be split.|
| _Destination_|Optional| **Variant**|A [Range](range-object-excel.md) object that represents the upper-left corner of the destination range for the parsed data. If this argument is omitted, Microsoft Excel parses in place.|

### Return Value

Variant


## Remarks

" `[xxx][xxx]`" would insert the first three characters into the first column of the destination range, and it would insert the next three characters into the second column. If this argument is omitted, Microsoft Excel guesses where to split the columns based on the spacing of the top left cell in the range. If you want to use a different range to guess the parse line, use a  **Range** object as the _ParseLine_ argument. That range must be one of the cells that's being parsed. The _ParseLine_ argument cannot be longer than 255 characters, including the brackets and spaces.


## Example

This example divides telephone numbers of the form 206-555-1212 into two columns. The first column contains only the area code, and the second column contains the seven-digit telephone number with the embedded hyphen.


```vb
Worksheets("Sheet1").Columns("A").Parse _ 
 parseLine:="[xxx] [xxxxxxxx]", _ 
 destination:=Worksheets("Sheet1").Range("B1")
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

