---
title: Rows.ConvertToText Method (Word)
keywords: vbawd10.chm155975890
f1_keywords:
- vbawd10.chm155975890
ms.prod: word
api_name:
- Word.Rows.ConvertToText
ms.assetid: 36b763f2-807b-53c0-e7ff-42b63bd356dc
ms.date: 06/08/2017
---


# Rows.ConvertToText Method (Word)

Converts rows in a table to text and returns a  **Range** object that represents the delimited text.


## Syntax

 _expression_ . **ConvertToText**( **_Separator_** , **_NestedTables_** )

 _expression_ Required. A variable that represents a **[Rows](rows-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Separator_|Optional| **Variant**|The character that delimits the converted columns (paragraph marks delimit the converted rows). Can be any following  **WdTableFieldSeparator** constants: **wdSeparateByCommas** , **wdSeparateByDefaultListSeparator** , **wdSeparateByParagraphs** , or **wdSeparateByTabs** (Default).|
| _NestedTables_|Optional| **Variant**| **True** if nested tables are converted to text. This argument is ignored if Separator is not **wdSeparateByParagraphs** . The default value is **True** .|

## See also


#### Concepts


[Rows Collection Object](rows-object-word.md)

