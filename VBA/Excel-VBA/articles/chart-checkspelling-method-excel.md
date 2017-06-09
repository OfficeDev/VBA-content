---
title: Chart.CheckSpelling Method (Excel)
keywords: vbaxl10.chm149092
f1_keywords:
- vbaxl10.chm149092
ms.prod: excel
api_name:
- Excel.Chart.CheckSpelling
ms.assetid: 9c39b0f1-4401-1399-58fa-444c9fa9fab4
ms.date: 06/08/2017
---


# Chart.CheckSpelling Method (Excel)

Checks the spelling of an object.


## Syntax

 _expression_ . **CheckSpelling**( **_CustomDictionary_** , **_IgnoreUppercase_** , **_AlwaysSuggest_** , **_SpellLang_** )

 _expression_ A variable that represents a **Chart** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _CustomDictionary_|Optional| **Variant**|A string that indicates the file name of the custom dictionary to be examined if the word isn't found in the main dictionary. If this argument is omitted, the currently specified dictionary is used.|
| _IgnoreUppercase_|Optional| **Variant**| **True** to have Microsoft Excel ignore words that are all uppercase. **False** to have Microsoft Excel check words that are all uppercase. If this argument is omitted, the current setting will be used.|
| _AlwaysSuggest_|Optional| **Variant**| **True** to have Microsoft Excel display a list of suggested alternate spellings when an incorrect spelling is found. **False** to have Microsoft Excel wait for you to input the correct spelling. If this argument is omitted, the current setting will be used.|
| _SpellLang_|Optional| **Variant**|The language of the dictionary being used. Can be one of the  **[MsoLanguageID](http://msdn.microsoft.com/library/65ea40f0-9a09-3d76-1519-4acddcc5f367%28Office.15%29.aspx)** values.|

## Remarks

This method has no return value; Microsoft Excel displays the  **Spelling** dialog box.

To check headers, footers, and objects on a worksheet, use this method on a  **[Worksheet](worksheet-object-excel.md)** object.


## See also


#### Concepts


[Chart Object](chart-object-excel.md)

