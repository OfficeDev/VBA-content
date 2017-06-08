---
title: Table.Sort Method (Word)
keywords: vbawd10.chm156303383
f1_keywords:
- vbawd10.chm156303383
ms.prod: word
api_name:
- Word.Table.Sort
ms.assetid: 2c68f7ad-2d57-05ea-bd8b-cb8712c21f02
ms.date: 06/08/2017
---


# Table.Sort Method (Word)

Sorts the specified table.


## Syntax

 _expression_ . **Sort**( **_ExcludeHeader_** , **_FieldNumber_** , **_SortFieldType_** , **_SortOrder_** , **_FieldNumber2_** , **_SortFieldType2_** , **_SortOrder2_** , **_FieldNumber3_** , **_SortFieldType3_** , **_SortOrder3_** , **_CaseSensitive_** , **_BidiSort_** , **_IgnoreThe_** , **_IgnoreKashida_** , **_IgnoreDiacritics_** , **_IgnoreHe_** , **_LanguageID_** )

 _expression_ Required. A variable that represents a **[Table](table-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ExcludeHeader_|Optional| **Variant**| **True** to exclude the first row from the sort operation. The default value is **False** .|
| _FieldNumber_|Optional| **Variant**|The first field by which to sort. Microsoft Word sorts by FieldNumber, then by FieldNumber2, and then by FieldNumber3.|
| _SortFieldType_|Optional| **Variant**|The sort type for FieldNumber. Can be one of the  **WdSortFieldType** constants. Some of these constants may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed. The default value is **wdSortFieldAlphanumeric** .|
| _SortOrder_|Optional| **Variant**|The sorting order to use when sorting FieldNumber. Can be a  **WdSortOrder** constant.|
| _FieldNumber2_|Optional| **Variant**|The second field by which to sort.|
| _SortFieldType2_|Optional| **Variant**|The sort type for FieldNumber2. Can be one of the  **WdSortFieldType** constants. Some of these constants may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed. The default value is **wdSortFieldAlphanumeric** .|
| _SortOrder2_|Optional| **Variant**|The sorting order to use when sorting FieldNumber2. Can be one  **WdSortOrder** constant.|
| _FieldNumber3_|Optional| **Variant**|The third field by which to sort.|
| _SortFieldType3_|Optional| **Variant**|The sort type for FieldNumber3. Can be one of the  **WdSortFieldType** constants. Some of these constants may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed. The default value is **wdSortFieldAlphanumeric** .|
| _SortOrder3_|Optional| **Variant**|The sorting order to use when sorting FieldNumber3. Can be one  **WdSortOrder** constant.|
| _CaseSensitive_|Optional| **Variant**| **True** to sort with case sensitivity. The default value is **False** .|
| _BidiSort_|Optional| **Variant**| **True** to sort based on right-to-left language rules. This argument may not be available to you, depending on the language support (U.S. English, for example) that you?ve selected or installed.|
| _IgnoreThe_|Optional| **Variant**| **True** to ignore the Arabic character alef lam when sorting right-to-left language text. This argument may not be available to you, depending on the language support (U.S. English, for example) that you?ve selected or installed.|
| _IgnoreKashida_|Optional| **Variant**| **True** to ignore kashidas when sorting right-to-left language text. This argument may not be available to you, depending on the language support (U.S. English, for example) that you?ve selected or installed.|
| _IgnoreDiacritics_|Optional| **Variant**| **True** to ignore bidirectional control characters when sorting right-to-left language text. This argument may not be available to you, depending on the language support (U.S. English, for example) that you?ve selected or installed.|
| _IgnoreHe_|Optional| **Variant**| **True** to ignore the Hebrew character he when sorting right-to-left language text. This argument may not be available to you, depending on the language support (U.S. English, for example) that you?ve selected or installed.|
| _LanguageID_|Optional| **Variant**|Specifies the sorting language. Can be one of the  **WdLanguageID** constants. Refer to the Object Browser for a list of the **WdLanguageID** constants.|

## Example

This example sorts the first table in the active document, excluding the heading row.


```vb
Sub NewTableSort() 
 ActiveDocument.Tables(1).Sort ExcludeHeader:=True 
End Sub
```


## See also


#### Concepts


[Table Object](table-object-word.md)

