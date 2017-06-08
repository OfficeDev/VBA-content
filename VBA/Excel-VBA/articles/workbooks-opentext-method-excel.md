---
title: Workbooks.OpenText Method (Excel)
keywords: vbaxl10.chm203083
f1_keywords:
- vbaxl10.chm203083
ms.prod: excel
api_name:
- Excel.Workbooks.OpenText
ms.assetid: a0771773-d0e2-13a0-e62b-51143e3f6bb1
ms.date: 06/08/2017
---


# Workbooks.OpenText Method (Excel)

Loads and parses a text file as a new workbook with a single sheet that contains the parsed text-file data.


## Syntax

 _expression_ . **OpenText**( **_Filename_** , **_Origin_** , **_StartRow_** , **_DataType_** , **_TextQualifier_** , **_ConsecutiveDelimiter_** , **_Tab_** , **_Semicolon_** , **_Comma_** , **_Space_** , **_Other_** , **_OtherChar_** , **_FieldInfo_** , **_TextVisualLayout_** , **_DecimalSeparator_** , **_ThousandsSeparator_** , **_TrailingMinusNumbers_** , **_Local_** )

 _expression_ A variable that represents a **Workbooks** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Filename_|Required| **String**|Specifies the file name of the text file to be opened and parsed.|
| _Origin_|Optional| **Variant**|Specifies the origin of the text file. Can be one of the following  **XlPlatform** constants: **xlMacintosh** , **xlWindows** , or **xlMSDOS** . Additionally, this could be an integer representing the code page number of the desired code page. For example, "1256" would specify that the encoding of the source text file is Arabic (Windows). If this argument is omitted, the method uses the current setting of the **File Origin** option in the **Text Import Wizard**.|
| _StartRow_|Optional| **Variant**|The row number at which to start parsing text. The default value is 1.|
| _DataType_|Optional| **Variant**| Specifies the column format of the data in the file. Can be one of the following **[XlTextParsingType](xltextparsingtype-enumeration-excel.md)** constants: **xlDelimited** or **xlFixedWidth** . If this argument is not specified, Microsoft Excel attempts to determine the column format when it opens the file.|
| _TextQualifier_|Optional| **[XlTextQualifier](xltextqualifier-enumeration-excel.md)**|Specifies the text qualifier.|
| _ConsecutiveDelimiter_|Optional| **Variant**| **True** to have consecutive delimiters considered one delimiter. The default is **False** .|
| _Tab_|Optional| **Variant**| **True** to have the tab character be the delimiter (DataType must be **xlDelimited** ). The default value is **False** .|
| _Semicolon_|Optional| **Variant**| **True** to have the semicolon character be the delimiter (DataType must be **xlDelimited** ). The default value is **False** .|
| _Comma_|Optional| **Variant**| **True** to have the comma character be the delimiter (DataType must be **xlDelimited** ). The default value is **False** .|
| _Space_|Optional| **Variant**| **True** to have the space character be the delimiter (DataType must be **xlDelimited** ). The default value is **False** .|
| _Other_|Optional| **Variant**| **True** to have the character specified by the OtherChar argument be the delimiter (DataType must be **xlDelimited** ). The default value is **False** .|
| _OtherChar_|Optional| **Variant**|(required if Other is  **True** ). Specifies the delimiter character when Other is **True** . If more than one character is specified, only the first character of the string is used; the remaining characters are ignored.|
| _FieldInfo_|Optional| **Variant**|An array containing parse information for individual columns of data. The interpretation depends on the value of DataType. When the data is delimited, this argument is an array of two-element arrays, with each two-element array specifying the conversion options for a particular column. The first element is the column number (1-based), and the second element is one of the  **[XlColumnDataType](xlcolumndatatype-enumeration-excel.md)** constants specifying how the column is parsed.|
| _TextVisualLayout_|Optional| **Variant**|The visual layout of the text.|
| _DecimalSeparator_|Optional| **Variant**|The decimal separator that Microsoft Excel uses when recognizing numbers. The default setting is the system setting.|
| _ThousandsSeparator_|Optional| **Variant**|The thousands separator that Excel uses when recognizing numbers. The default setting is the system setting.|
| _TrailingMinusNumbers_|Optional| **Variant**|Specify  **True** if numbers with a minus character at the end should be treated as negative numbers. If **False** or omitted, numbers with a minus character at the end are treated as text.|
| _Local_|Optional| **Variant**|Specify  **True** if regional settings of the machine should be used for separators, numbers and data formatting.|

## Remarks

 **_FieldInfo_ Parameter Information**

You can use  **xlEMDFormat** only if you have installed and selected Taiwanese language support. The **xlEMDFormat** constant specifies that Taiwanese era dates are being used.

The column specifiers can be in any order. If there's no column specifier for a particular column in the input data, the column is parsed with the General setting.


|**Note**|
|:-----|  
|This example causes the third column to be parsed as MDY (for example, 01/10/1970), the first column to be parsed as text, and the remaining columns in the source data to be parsed with the General setting.|

```vb
Array(Array(3, 3), Array(1, 2))
```

If the source data has fixed-width columns, the first element in each two-element array specifies the position of the starting character in the column (as an integer; character 0 (zero) is the first character). The second element in the two-element array specifies the parse option for the column as a number between 0 and 9, as listed in the preceding table.

 **_ThousandsSeparator_ Parameter Information**

The following table shows the results of importing text into Excel for various import settings. Numeric results are displayed in the rightmost column.



|**System decimal separator**|**System thousands separator**|**Decimal separator value**|**Thousands separator value**|**Text imported**|**Cell value (data type)**|
|:-----|:-----|:-----|:-----|:-----|:-----|
|Period|Comma|Comma|Period|123.123,45|123,123.45 (numeric)|
|Period|Comma|Comma|Comma|123.123,45|123.123,45 (text)|
|Comma|Period|Period|Comma|123,123.45|123,123.45 (numeric)|
|Period|Comma|Period|Comma|123 123.45|123 123.45 (text)|
|Period|Comma|Period|Space|123 123.45|123,123.45 (numeric)|

## Example

This example opens the file Data.txt and uses tab delimiters to parse the text file into a worksheet.


```vb
Workbooks.OpenText filename:="DATA.TXT", _ 
    dataType:=xlDelimited, tab:=True
```


## See also


#### Concepts


[Workbooks Object](workbooks-object-excel.md)

