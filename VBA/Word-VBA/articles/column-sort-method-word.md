---
title: Column.Sort Method (Word)
keywords: vbawd10.chm156172492
f1_keywords:
- vbawd10.chm156172492
ms.prod: word
api_name:
- Word.Column.Sort
ms.assetid: c71dbc55-a0dc-4ced-f579-6b446c427f88
ms.date: 06/08/2017
---


# Column.Sort Method (Word)

Sorts the specified table column.


## Syntax

 _expression_ . **Sort**( **_ExcludeHeader_** , **_SortFieldType_** , **_SortOrder_** , **_CaseSensitive_** , **_BidiSort_** , **_IgnoreThe_** , **_IgnoreKashida_** , **_IgnoreDiacritics_** , **_IgnoreHe_** , **_LanguageID_** )

 _expression_ Required. A variable that represents a **[Column](column-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ExcludeHeader_|Optional| **Variant**| **True** to exclude the first row or paragraph header from the sort operation. The default value is **False** .|
| _SortFieldType_|Optional| **Variant**|The sort type for the column. Can be one of the  **[WdSortFieldType](wdsortfieldtype-enumeration-word.md)** constants.|
| _SortOrder_|Optional| **Variant**|The sorting order to use for the column. Can be one  **[WdSortOrder](wdsortorder-enumeration-word.md)** constant.|
| _CaseSensitive_|Optional| **Variant**| **True** to sort with case sensitivity. The default value is **False** .|
| _BidiSort_|Optional| **Variant**| **True** to sort based on right-to-left language rules. This argument may not be available to you, depending on the language support (U.S. English, for example) that you?ve selected or installed.|
| _IgnoreThe_|Optional| **Variant**| **True** to ignore the Arabic character alef lam when sorting right-to-left language text. This argument may not be available to you, depending on the language support (U.S. English, for example) that you?ve selected or installed.|
| _IgnoreKashida_|Optional| **Variant**| **True** to ignore kashidas when sorting right-to-left language text. This argument may not be available to you, depending on the language support (U.S. English, for example) that you?ve selected or installed.|
| _IgnoreDiacritics_|Optional| **Variant**| **True** to ignore bidirectional control characters when sorting right-to-left language text. This argument may not be available to you, depending on the language support (U.S. English, for example) that you?ve selected or installed.|
| _IgnoreHe_|Optional| **Variant**| **True** to ignore the Hebrew character he when sorting right-to-left language text. This argument may not be available to you, depending on the language support (U.S. English, for example) that you?ve selected or installed.|
| _LanguageID_|Optional| **Variant**|Specifies the sorting language. Can be one of the  **[WdLanguageID](wdlanguageid-enumeration-word.md)** constants.|

## Remarks

If you want to sort paragraphs within a table cell, include only the paragraphs and not the end-of-cell mark; if you include the end-of-cell mark in a selection or range and then attempt to sort the paragraphs, Word displays a message stating that it found no valid records to sort.


## See also


#### Concepts


[Column Object](column-object-word.md)

