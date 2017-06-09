---
title: Indexes.Add Method (Word)
keywords: vbawd10.chm159121512
f1_keywords:
- vbawd10.chm159121512
ms.prod: word
api_name:
- Word.Indexes.Add
ms.assetid: 324fe848-5c93-d456-0e45-42be116b1b90
ms.date: 06/08/2017
---


# Indexes.Add Method (Word)

Returns an  **Index** object that represents a new index added to a document.


## Syntax

 _expression_ . **Add**( **_Range_** , **_HeadingSeparator_** , **_RightAlignPageNumbers_** , **_Type_** , **_NumberOfColumns_** , **_AccentedLetters_** , **_SortBy_** , **_IndexLanguage_** )

 _expression_ Required. A variable that represents an **[Indexes](indexes-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **Range**|The range where you want the index to appear. The index replaces the range, if the range is not collapsed.|
| _HeadingSeparator_|Optional| **Variant**|The text between alphabetical groups (entries that start with the same letter) in the index. Can be one of the  **WdHeadingSeparator** constants.|
| _RightAlignPageNumbers_|Optional| **Variant**| **True** to align page numbers with the right margin.|
| _Type_|Optional| **Variant**|Specifies whether subentries are on the same line (run-in) as the main entry or on a separate line (indented) from the main entry. Can be either of the following  **WdIndexType** constants: **wdIndexIndent** or **wdIndexRunin** .|
| _NumberOfColumns_|Optional| **Variant**|The number of columns for each page of the index. Specifying 0 (zero) sets the number of columns in the index to the same number as in the document.|
| _AccentedLetters_|Optional| **Variant**| **True** to include separate headings for accented letters in the index (for example, words that begin with "?" and words that begin with "A" are listed under separate headings).|
| _SortBy_|Optional| **Variant**|The sorting criteria to be used for the specified index. Can be either of the following  **WdIndexSortBy** constants: **wdIndexSortByStroke** or **wdIndexSortBySyllable** .|
| _IndexLanguage_|Optional| **Variant**|The sorting language to be used for the specified index. Can be any of the  **WdLanguageID** constants. For the list of valid **WdLanguageID** constants, see the Object Browser in the Visual Basic Editor.|

### Return Value

Index


## Remarks

An index is built from Index Entry (XE) fields in a document. Use the  **MarkEntry** method to mark index entries to be included in an index.


## Example

This example marks an index entry, and then it creates an index at the end of the active document.


```vb
ActiveDocument.Indexes.MarkEntry _ 
 Range:=Selection.Range, Entry:="My Entry" 
Set MyRange = ActiveDocument.Content 
MyRange.Collapse Direction:=wdCollapseEnd 
ActiveDocument.Indexes.Add Range:=MyRange, Type:=wdIndexRunin
```


## See also


#### Concepts


[Indexes Collection Object](indexes-object-word.md)

