---
title: Range.Sort Method (Word)
keywords: vbawd10.chm157155812
f1_keywords:
- vbawd10.chm157155812
ms.prod: word
api_name:
- Word.Range.Sort
ms.assetid: 2030f99e-0307-d2b7-9e14-1d0888f3fda6
ms.date: 06/08/2017
---


# Range.Sort Method (Word)

Sorts the paragraphs in the specified range.


## Syntax

 _expression_ . **Sort**( **_ExcludeHeader_** , **_FieldNumber_** , **_SortFieldType_** , **_SortOrder_** , **_FieldNumber2_** , **_SortFieldType2_** , **_SortOrder2_** , **_FieldNumber3_** , **_SortFieldType3_** , **_SortOrder3_** , **_SortColumn_** , **_Separator_** , **_CaseSensitive_** , **_BidiSort_** , **_IgnoreThe_** , **_IgnoreKashida_** , **_IgnoreDiacritics_** , **_IgnoreHe_** , **_LanguageID_** )

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ExcludeHeader_|Optional| **Variant**| **True** to exclude the first row or paragraph header from the sort operation. The default value is **False** .|
| _FieldNumber_|Optional| **Variant**|The fields by which to sort. Microsoft Word sorts by FieldNumber, then by FieldNumber2, and then by FieldNumber3.|
| _SortFieldType_|Optional| **Variant**|The respective sort types for FieldNumber. Can be one of the  **WdSortFieldType** constants. The default value is **wdSortFieldAlphanumeric** . Some of these constants may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.|
| _SortOrder_|Optional| **Variant**|The sorting order to use when sorting FieldNumber. Can be any  **WdSortOrder** constant.|
| _FieldNumber2_|Optional| **Variant**|The fields by which to sort.|
| _SortFieldType2_|Optional| **Variant**|The respective sort types for FieldNumber2. Can be one of the  **WdSortFieldType** constants. The default value is **wdSortFieldAlphanumeric** . Some of these constants may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.|
| _SortOrder2_|Optional| **Variant**|The sorting order to use when sorting FieldNumber2. Can be any  **WdSortOrder** constant.|
| _FieldNumber3_|Optional| **Variant**|The fields by which to sort.|
| _SortFieldType3_|Required||Some of these constants may not be available to you, depending on the language support (U.S. English, for example) that you?ve selected or installed. The default value is  **wdSortFieldAlphanumeric** . Some of these constants may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.|
| _SortOrder3_|Optional| **Variant**|The sorting order to use when sorting FieldNumber3. Can be any  **WdSortOrder** constant.|
| _SortColumn_|Optional| **Variant**| **True** to sort only the column specified by the **Range** object.|
| _Separator_|Optional| **Variant**|The type of field separator. Can be one of the  **WdSortSeparator** constants.|
| _CaseSensitive_|Optional| **Variant**| **True** to sort with case sensitivity. The default value is **False** .|
| _BidiSort_|Optional| **Variant**| **True** to sort based on right-to-left language rules. This argument may not be available to you, depending on the language support (U.S. English, for example) that you?ve selected or installed.|
| _IgnoreThe_|Optional| **Variant**| **True** to ignore the Arabic character alef lam when sorting right-to-left language text. This argument may not be available to you, depending on the language support (U.S. English, for example) that you?ve selected or installed.|
| _IgnoreKashida_|Optional| **Variant**| **True** to ignore kashidas when sorting right-to-left language text. This argument may not be available to you, depending on the language support (U.S. English, for example) that you?ve selected or installed.|
| _IgnoreDiacritics_|Optional| **Variant**| **True** to ignore bidirectional control characters when sorting right-to-left language text. This argument may not be available to you, depending on the language support (U.S. English, for example) that you?ve selected or installed.|
| _IgnoreHe_|Optional| **Variant**| **True** to ignore the Hebrew character he when sorting right-to-left language text. This argument may not be available to you, depending on the language support (U.S. English, for example) that you?ve selected or installed.|
| _LanguageID_|Optional| **Variant**|Specifies the sorting language. Can be one of the  **WdLanguageID** constants. Refer to the Object Browser for a list of the **WdLanguageID** constants.|

## Example

This example inserts three lines of text into a new document and then sorts the lines in ascending alphanumeric order


```vb
Sub NewParagraphSort() 
 Dim newDoc As Document 
 Set newDoc = Documents.Add 
 newDoc.Content.InsertAfter "pear" &; Chr(13) _ 
 &; "zucchini" &; Chr(13) &; "apple" &; Chr(13) 
 newDoc.Content.Sort SortOrder:=wdSortOrderAscending 
End Sub
```


## See also


#### Concepts


[Range Object](range-object-word.md)

