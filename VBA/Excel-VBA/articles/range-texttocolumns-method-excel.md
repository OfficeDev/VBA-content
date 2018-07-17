---
title: Range.TextToColumns Method (Excel)
keywords: vbaxl10.chm144210
f1_keywords:
- vbaxl10.chm144210
ms.prod: excel
api_name:
- Excel.Range.TextToColumns
ms.assetid: 0b0bf329-ab99-7edc-1b8f-aad03513abde
ms.date: 06/08/2017
---


# Range.TextToColumns Method (Excel)

Parses a column of cells that contain text into several columns.


## Syntax

 _expression_ . **TextToColumns**( **_Destination_** , **_DataType_** , **_TextQualifier_** , **_ConsecutiveDelimiter_** , **_Tab_** , **_Semicolon_** , **_Comma_** , **_Space_** , **_Other_** , **_OtherChar_** , **_FieldInfo_** , **_DecimalSeparator_** , **_ThousandsSeparator_** , **_TrailingMinusNumbers_** )

 _expression_ A variable that represents a **Range** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Destination_|Optional| **Variant**|A  **Range** object that specifies where Microsoft Excel will place the results. If the range is larger than a single cell, the top left cell is used.|
| _DataType_|Optional| **[XlTextParsingType](xltextparsingtype-enumeration-excel.md)**|The format of the text to be split into columns.|
| _TextQualifier_|Optional| **[XlTextQualifier](xltextqualifier-enumeration-excel.md)**|Specifies whether to use single, double or no quotes as the text qualifier.|
| _ConsecutiveDelimiter_|Optional| **Variant**| **True** to have Microsoft Excel consider consecutive delimiters as one delimiter. The default value is **False** .|
| _Tab_|Optional| **Variant**| **True** to have _DataType_ be **xlDelimited** and to have the tab character be a delimiter. The default value is **False** .|
| _Semicolon_|Optional| **Variant**| **True** to have _DataType_ be **xlDelimited** and to have the semicolon be a delimiter. The default value is **False** .|
| _Comma_|Optional| **Variant**| **True** to have _DataType_ be **xlDelimited** and to have the comma be a delimiter. The default value is **False** .|
| _Space_|Optional| **Variant**| **True** to have _DataType_ be **xlDelimited** and to have the space character be a delimiter. The default value is **False** .|
| _Other_|Optional| **Variant**| **True** to have _DataType_ be **xlDelimited** and to have the character specified by the _OtherChar_ argument be a delimiter. The default value is **False** .|
| _OtherChar_|Optional| **Variant**|(required if  _Other_ is **True** ). The delimiter character when _Other_ is **True** . If more than one character is specified, only the first character of the string is used; the remaining characters are ignored.|
| _FieldInfo_|Optional| **Variant**|An array containing parse information for the individual columns of data. The interpretation depends on the value of  _DataType_, but generally, this argument is an array of two-element arrays, each one specifying the conversion options for a particular column. The first element is an integer, and the second element is one of the [xlColumnDataType](xlcolumndatatype-enumeration-excel.md) constants. See Remarks for more detail.|
| _DecimalSeparator_|Optional| **Variant**|The decimal separator that Microsoft Excel uses when recognizing numbers. The default setting is the system setting.|
| _ThousandsSeparator_|Optional| **Variant**|The thousands separator that Excel uses when recognizing numbers. The default setting is the system setting.|
| _TrailingMinusNumbers_|Optional| **Variant**|Numbers that begin with a minus character.|

### Return Value

Variant


## Remarks

The following table shows the results of importing text into Excel for various import settings. Numeric results are displayed in the rightmost column.



|**System decimal separator**|**System thousands separator**|**Decimal separator value**|**Thousands separator value**|**Original text**|**Cell value (data type)**|
|:-----|:-----|:-----|:-----|:-----|:-----|
|Period|Comma|Comma|Period|123.123,45|123,123.45 (numeric)|
|Period|Comma|Comma|Comma|123.123,45|123.123,45 (text)|
|Comma|Period|Comma|Period|123,123.45|123,123.45 (numeric)|
|Period|Comma|Period|Comma|123 123.45|123 123.45 (text)|
|Period|Comma|Period|Space|123 123.45|123,123.45 (numeric)|


### FieldInfo

_FieldInfo_ consists of an array of two-element arrays, with the first element as an integer representing either the column number or starting character position, and the second element as one of the [xlColumnDataType](xlcolumndatatype-enumeration-excel.md) constants that specifies the parse option for the column, as listed below.

#### First element

The interpretation of the first element depends on the value of  _DataType_.

|**DataType** equals|**First element**|
|:-----|:-----|
|xlDelimited|Column number (1-based)|
|xlFixedWidth|Starting character position for the column (0-based)|


If _DataType_ is **xlDelimited**, the first element is the column number (1-based). The two-element arrays can be in any order, regardless of column number. If a given column number is not present for a particular column in the input data, the column is parsed with the  **xlGeneralFormat** setting. This example causes the third column to be skipped, the first column to be parsed as text, and the remaining columns in the source data to be parsed with the **xlGeneralFormat** setting.

 `Array(Array(3, 9), Array(1, 2))`


If the source data has fixed-width columns, the first element specifies the starting character position as an integer for the column; 0 (zero) is the first character. The following example parses two columns from a fixed-width file, with the first column starting at the beginning of the line and extending for 10 characters. The second column starts at position 15 and goes to the end of the line. To avoid including the characters between position 10 and position 15, Microsoft Excel adds a skipped column entry.

 `Array(Array(0, 1), Array(10, 9), Array(15, 1))`

#### Second element

The second element of the two-element child arrays of _FieldInfo_ can be one of these **XlColumnDataType** constants.

|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
|**xlGeneralFormat**|1|General|
|**xlTextFormat**|2|Text|
|**xlMDYFormat**|3|MDY Date|
|**xlDMYFormat**|4|DMY Date|
|**xlYMDFormat**|5|YMD Date|
|**xlMYDFormat**|6|MYD Date|
|**xlDYMFormat**|7|DYM Date|
|**xlYDMFormat**|8|YDM Date|
|**xlEMDFormat**|10|EMD Date|
|**xlSkipColumn**|9|Skip Column|

You can use  **xlEMDFormat** only if Taiwanese language support is installed and selected. The **xlEMDFormat** constant specifies that Taiwanese era dates are being used.



## Example

This example converts the contents of the Clipboard, which contains a space-delimited text table, into separate columns on Sheet1. You can create a simple space-delimited table in Notepad or WordPad (or another text editor), copy the text table to the Clipboard, switch to Microsoft Excel, and then run this example.


```vb
Worksheets("Sheet1").Activate 
ActiveSheet.Paste 
Selection.TextToColumns DataType:=xlDelimited, _ 
 ConsecutiveDelimiter:=True, Space:=True
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

