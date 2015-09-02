
# Range.TextToColumns Method (Excel)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Parses a column of cells that contain text into several columns.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **TextToColumns**( **_Destination_**,  **_DataType_**,  **_TextQualifier_**,  **_ConsecutiveDelimiter_**,  **_Tab_**,  **_Semicolon_**,  **_Comma_**,  **_Space_**,  **_Other_**,  **_OtherChar_**,  **_FieldInfo_**,  **_DecimalSeparator_**,  **_ThousandsSeparator_**,  **_TrailingMinusNumbers_**)

 _expression_A variable that represents a  **Range** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Destination|Optional| **Variant**|A  **Range** object that specifies where Microsoft Excel will place the results. If the range is larger than a single cell, the top left cell is used.|
|DataType|Optional| ** [XlTextParsingType](71d76a41-c0b0-0b0f-27b5-7cac0d4c4ac4.md)**|The format of the text to be split into columns.|
|TextQualifier|Optional| ** [XlTextQualifier](ba209892-9dea-84db-eafd-629c7ab0b20f.md)**|Specifies whether to use single, double or no quotes as the text qualifier.|
|ConsecutiveDelimiter|Optional| **Variant**| **True** to have Microsoft Excel consider consecutive delimiters as one delimiter. The default value is **False**.|
|Tab|Optional| **Variant**| **True** to haveDataType be **xlDelimited** and to have the tab character be a delimiter. The default value is **False**.|
|Semicolon|Optional| **Variant**| **True** to haveDataType be **xlDelimited** and to have the semicolon be a delimiter. The default value is **False**.|
|Comma|Optional| **Variant**| **True** to haveDataType be **xlDelimited** and to have the comma be a delimiter. The default value is **False**.|
|Space|Optional| **Variant**| **True** to haveDataType be **xlDelimited** and to have the space character be a delimiter. The default value is **False**.|
|Other|Optional| **Variant**| **True** to haveDataType be **xlDelimited** and to have the character specified by theOtherChar argument be a delimiter. The default value is **False**.|
|OtherChar|Optional| **Variant**|(required if Other is **True**). The delimiter character when Other is **True**. If more than one character is specified, only the first character of the string is used; the remaining characters are ignored.|
|FieldInfo|Optional| **Variant**|An array containing parse information for the individual columns of data. The interpretation depends on the value of DataType. When the data is delimited, this argument is an array of two-element arrays, with each two-element array specifying the conversion options for a particular column. The first element is the column number (1-based), and the second element is one of the  [xlColumnDataType](034f6011-c860-0887-9661-857821f630e4.md)constants specifying how the column is parsed.|
|DecimalSeparator|Optional| **Variant**|The decimal separator that Microsoft Excel uses when recognizing numbers. The default setting is the system setting.|
|ThousandsSeparator|Optional| **Variant**|The thousands separator that Excel uses when recognizing numbers. The default setting is the system setting.|
|TrailingMinusNumbers|Optional| **Variant**|Numbers that begin with a minus character.|

### Return Value

Variant


## Remarks
<a name="sectionSection1"> </a>

The following table shows the results of importing text into Excel for various import settings. Numeric results are displayed in the rightmost column.



|**System decimal separator**|**System thousands separator**|**Decimal separator value**|**Thousands separator value**|**Original text**|**Cell value (data type)**|
|:-----|:-----|:-----|:-----|:-----|:-----|
|Period|Comma|Comma|Period|123.123,45|123,123.45 (numeric)|
|Period|Comma|Comma|Comma|123.123,45|123.123,45 (text)|
|Comma|Period|Comma|Period|123,123.45|123,123.45 (numeric)|
|Period|Comma|Period|Comma|123 123.45|123 123.45 (text)|
|Period|Comma|Period|Space|123 123.45|123,123.45 (numeric)|


| **XlColumnDataType** can be one of these **XlColumnDataType** constants.|
| **xlGeneralFormat**. General|
| **xlTextFormat**. Text|
| **xlMDYFormat**. MDY Date|
| **xlDMYFormat**. DMY Date|
| **xlYMDFormat**. YMD Date|
| **xlMYDFormat**. MYD Date|
| **xlDYMFormat**. DYM Date|
| **xlYDMFormat**. YDM Date|
| **xlEMDFormat**. EMD Date|
| **xlSkipColumn**. Skip Column|
You can use  **xlEMDFormat** only if Taiwanese language support is installed and selected. The **xlEMDFormat** constant specifies that Taiwanese era dates are being used.

The column specifiers can be in any order. If a given column specifier is not present for a particular column in the input data, the column is parsed with the  **xlGeneralFormat** setting. This example causes the third column to be skipped, the first column to be parsed as text, and the remaining columns in the source data to be parsed with the **xlGeneralFormat** setting.

 `Array(Array(3, 9), Array(1, 2))`

If the source data has fixed-width columns, the first element of each two-element array specifies the starting character position in the column (as an integer; 0 (zero) is the first character). The second element of the two-element array specifies the parse option for the column as a number from 1 through 9, as listed above.

The following example parses two columns from a fixed-width file, with the first column starting at the beginning of the line and extending for 10 characters. The second column starts at position 15 and goes to the end of the line. To avoid including the characters between position 10 and position 15, Microsoft Excel adds a skipped column entry.

 `Array(Array(0, 1), Array(10, 9), Array(15, 1))`


## Example
<a name="sectionSection2"> </a>

This example converts the contents of the Clipboard, which contains a space-delimited text table, into separate columns on Sheet1. You can create a simple space-delimited table in Notepad or WordPad (or another text editor), copy the text table to the Clipboard, switch to Microsoft Excel, and then run this example.


```
Worksheets("Sheet1").Activate 
ActiveSheet.Paste 
Selection.TextToColumns DataType:=xlDelimited, _ 
 ConsecutiveDelimiter:=True, Space:=True
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Range Object](b8207778-0dcc-4570-1234-f130532cc8cd.md)
#### Other resources


 [Range Object Members](4336bf81-1e63-7e44-1792-baf366a027a7.md)
