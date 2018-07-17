---
title: Row.ConvertToText Method (Word)
keywords: vbawd10.chm156237842
f1_keywords:
- vbawd10.chm156237842
ms.prod: word
api_name:
- Word.Row.ConvertToText
ms.assetid: ca26c76c-0695-58b9-7a7a-a74f3350a8e9
ms.date: 06/08/2017
---


# Row.ConvertToText Method (Word)

Converts a table to text and returns a  **Range** object that represents the delimited text.


## Syntax

 _expression_ . **ConvertToText**( **_Separator_** , **_NestedTables_** )

 _expression_ Required. A variable that represents a **[Row](row-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Separator_|Optional| **Variant**|The character that delimits the converted columns (paragraph marks delimit the converted rows). Can be any following  **WdTableFieldSeparator** constants: **wdSeparateByCommas** , **wdSeparateByDefaultListSeparator** , **wdSeparateByParagraphs** , or **wdSeparateByTabs** (Default).|
| _NestedTables_|Optional| **Variant**| **True** if nested tables are converted to text. This argument is ignored if Separator is not **wdSeparateByParagraphs** . The default value is **True** .|

## See also


#### Concepts


[Row Object](row-object-word.md)

