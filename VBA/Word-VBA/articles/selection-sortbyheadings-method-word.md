---
title: Selection.SortByHeadings Method (Word)
keywords: vbawd10.chm158663698
f1_keywords:
- vbawd10.chm158663698
ms.prod: word
ms.assetid: fc38c337-b658-7b8d-2191-2ee98a93b48e
ms.date: 06/08/2017
---


# Selection.SortByHeadings Method (Word)

Sorts the headings in the specified selection.


## Syntax

 _expression_ . **SortByHeadings**_(SortFieldType,_ _SortOrder,_ _CaseSensitive,_ _BidiSort,_ _IgnoreThe,_ _IgnoreKashida,_ _IgnoreDiacritics,_ _IgnoreHe,_ _LanguageID)_

 _expression_ A variable that represents a **Selection** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
|||||
| _SortFieldType_|Optional|VARIANT|The sort field type to use. Can be one of the [WdSortFieldType](wdsortfieldtype-enumeration-word.md) constants. The default value is **wdSortFieldAlphanumeric**. Depending on the language support (U.S. English, for example) that you have selected or installed, some of these constants may not be available to you.|
| _SortOrder_|Optional|VARIANT|The sorting order to use. Can be one of the [WdSortOrder](wdsortorder-enumeration-word.md) constants.|
| _CaseSensitive_|Optional|VARIANT| **True** to sort with case sensitivity. The default value is **False**.|
| _BidiSort_|Optional|VARIANT| **True** to sort based on right-to-left language rules. Depending on the language support (U.S. English, for example) that you have selected or installed, this parameter may not be available to you.|
| _IgnoreThe_|Optional|VARIANT| **True** to ignore the Arabic character alef lam when sorting right-to-left language text. Depending on the language support (U.S. English, for example) that you have selected or installed, this parameter may not be available to you.|
| _IgnoreKashida_|Optional|VARIANT| **True** to ignore kashidas when sorting right-to-left language text. Depending on the language support (U.S. English, for example) that you have selected or installed, this parameter may not be available to you.|
| _IgnoreDiacritics_|Optional|VARIANT| **True** to ignore bidirectional control characters when sorting right-to-left language text. Depending on the language support (U.S. English, for example) that you have selected or installed, this parameter may not be available to you.|
| _IgnoreHe_|Optional|VARIANT| **True** to ignore the Hebrew character he when sorting right-to-left language text. Depending on the language support (U.S. English, for example) that you have selected or installed, this parameter may not be available to you.|
| _LanguageID_|Optional|VARIANT|Specifies the sorting language. Can be one of the [WdLanguageID](wdlanguageid-enumeration-word.md) constants.|

### Return value

 **VOID**


## See also


#### Concepts


[Selection Object](selection-object-word.md)

