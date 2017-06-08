---
title: Selection.Sort Method (Word)
keywords: vbawd10.chm158663679
f1_keywords:
- vbawd10.chm158663679
ms.prod: word
api_name:
- Word.Selection.Sort
ms.assetid: 3f29f6bf-a6b4-1638-b078-f61a4f36c17e
ms.date: 06/08/2017
---


# Selection.Sort Method (Word)

Sorts the paragraphs in the specified selection.


## Syntax

 _expression_ . **Sort**( **_ExcludeHeader_** , **_FieldNumber_** , **_SortFieldType_** , **_SortOrder_** , **_FieldNumber2_** , **_SortFieldType2_** , **_SortOrder2_** , **_FieldNumber3_** , **_SortFieldType3_** , **_SortOrder3_** , **_SortColumn_** , **_Separator_** , **_CaseSensitive_** , **_BidiSort_** , **_IgnoreThe_** , **_IgnoreKashida_** , **_IgnoreDiacritics_** , **_IgnoreHe_** , **_LanguageID_** , **_SubFieldNumber_** , **_SubFieldNumber2_** , **_SubFieldNumber3_** )

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ExcludeHeader_|Optional| **Variant**| **True** to exclude the first row or paragraph header from the sort operation. The default value is **False** .|
| _FieldNumber_|Optional| **Variant**|The first field by which to sort.|
| _SortFieldType_|Optional| **Variant**|The sort type for FieldNumber. Can be one of the  **WdSortFieldType** constants. The default value is **wdSortFieldAlphanumeric** . Some of the **WdSortFieldType** constants may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.|
| _SortOrder_|Optional| **Variant**|The sorting order to use when sorting FieldNumber. Can be one  **[WdSortOrder](wdsortorder-enumeration-word.md)** constant.The default value is **wdSortOrderAscending** .|
| _FieldNumber2_|Optional| **Variant**|The second field by which to sort.|
| _SortFieldType2_|Optional| **Variant**|The sort type for FieldNumber2. Can be one of the  **WdSortFieldType** constants. The default value is **wdSortFieldAlphanumeric** . Some of the **WdSortFieldType** constants may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.|
| _SortOrder_|Optional| **Variant**|The sorting order to use when sorting FieldNumber2. Can be one  **[WdSortOrder](wdsortorder-enumeration-word.md)** constant.The default value is **wdSortOrderAscending** .|
| _SortColumn2_|Optional| **Variant**| **True** to sort only the column specified by the **Selection** object.|
| _Separator_|Optional| **Variant**|The type of field separator.|
| _FieldNumber3_|Optional| **Variant**|The third field by which to sort.|
| _SortFieldType3_|Optional| **Variant**|The sort type for FieldNumber3. Can be one of the  **WdSortFieldType** constants. The default value is **wdSortFieldAlphanumeric** . Some of the **WdSortFieldType** constants may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.|
| _SortOrder3_|Optional| **Variant**|The sorting order to use when sorting FieldNumber3. Can be one  **[WdSortOrder](wdsortorder-enumeration-word.md)** constant.The default value is **wdSortOrderAscending** .|
| _CaseSensitive_|Optional| **Variant**| **True** to sort with case sensitivity. The default value is **False** .|
| _BidiSort_|Optional| **Variant**| **True** to sort based on right-to-left language rules. This argument may not be available to you, depending on the language support (U.S. English, for example) that you?ve selected or installed.|
| _IgnoreThe_|Optional| **Variant**| **True** to ignore the Arabic character alef lam when sorting right-to-left language text. This argument may not be available to you, depending on the language support (U.S. English, for example) that you?ve selected or installed.|
| _IgnoreKashida_|Optional| **Variant**| **True** to ignore kashidas when sorting right-to-left language text. This argument may not be available to you, depending on the language support (U.S. English, for example) that you?ve selected or installed.|
| _IgnoreDiacritics_|Optional| **Variant**| **True** to ignore bidirectional control characters when sorting right-to-left language text. This argument may not be available to you, depending on the language support (U.S. English, for example) that you?ve selected or installed.|
| _IgnoreHe_|Optional| **Variant**| **True** to ignore the Hebrew character he when sorting right-to-left language text. This argument may not be available to you, depending on the language support (U.S. English, for example) that you?ve selected or installed.|
| _LanguageID_|Optional| **Variant**|Specifies the sorting language. Can be one of the  **WdLanguageID** constants.|
| _SubFieldNumber_|Optional| **Variant**|A secondary field number by which to sort.|
| _SubFieldNumber2_|Optional| **Variant**|A secondary field number by which to sort.|
| _SubFieldNumber3_|Optional| **Variant**|A secondary field number by which to sort.|

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


[Selection Object](selection-object-word.md)

