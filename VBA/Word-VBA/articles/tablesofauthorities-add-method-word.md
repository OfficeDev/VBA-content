---
title: TablesOfAuthorities.Add Method (Word)
keywords: vbawd10.chm152174692
f1_keywords:
- vbawd10.chm152174692
ms.prod: word
api_name:
- Word.TablesOfAuthorities.Add
ms.assetid: 8d89d4cd-933e-eb54-5644-fe02c81fb4a1
ms.date: 06/08/2017
---


# TablesOfAuthorities.Add Method (Word)

Returns a  **TableOfAuthorities** object that represents a table of authorities added to a document.


## Syntax

 _expression_ . **Add**( **_Range_** , **_Category_** , **_Bookmark_** , **_Passim_** , **_KeepEntryFormatting_** , **_Separator_** , **_IncludeSequenceName_** , **_EntrySeparator_** , **_PageRangeSeparator_** , **_IncludeCategoryHeader_** , **_PageNumberSeparator_** )

 _expression_ Required. A variable that represents a **[TablesOfAuthorities](tablesofauthorities-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **Range object**|The range where you want the table of authorities to appear. The table of authorities replaces the range, if the range isn't collapsed.|
| _Category_|Optional| **Variant**|The category of entries you want to include in the table of authorities. Corresponds to the \c switch for a Table of Authorities (TOA) field. Values 0 through 16 correspond to the items listed in the  **Category** box on the **Table of Authorities** tab in the **Index and Tables** dialog box ( **Reference** command, **Insert** menu). The default value is 1.|
| _Bookmark_|Optional| **Variant**|The string name of the bookmark from which you want to collect entries for the table of authorities. If Bookmark is specified, the entries are collected only from the portion of the document marked by the bookmark. Corresponds to the \b switch for a Table of Authorities (TOA) field.|
| _Passim_|Optional| **Variant**| **True** to replace five or more page references to the same authority with Passim in the table of authorities. Corresponds to the \p switch for a Table of Authorities (TOA) field. If this argument is omitted, Passim is assumed to be **False** .|
| _KeepEntryFormatting_|Optional| **Variant**| **True** to apply formatting from table of authorities entries to the entries in the table of authorities. Corresponds to the \f switch for a Table of Authorities (TOA) field. If this argument is omitted, KeepEntryFormatting is assumed to be **True** .|
| _Separator_|Optional| **Variant**|The characters (up to five) between each sequence number and its page number in the table of authorities. Corresponds to the \d switch for a Table of Authorities (TOA) field. If argument is omitted, a hyphen (-) is used.|
| _IncludeSequenceName_|Optional| **Variant**|A string that specifies the Sequence (SEQ) field identifier for the table of authorities. Corresponds to the \s switch for a Table of Authorities (TOA) field.|
| _EntrySeparator_|Optional| **Variant**|The characters (up to five) that separate each entry and its page number in the table of authorities. Corresponds to the \e switch for a Table of Authorities (TOA) field. If this argument is omitted, no separator is used.|
| _PageRangeSeparator_|Optional| **Variant**|The characters (up to five) that separate the beginning and ending page numbers in each page range the table of authorities. Corresponds to the \g switch for a Table of Authorities (TOA) field. If this argument is omitted, an en dash is used.|
| _IncludeCategoryHeader_|Optional| **Variant**| **True** to have the category name for each group of entries appear in the table of authorities (for example, Cases). Corresponds to the \h switch for a Table of Authorities (TOA) field. If this argument is omitted, IncludeCategoryHeader is assumed to be **True** .|
| _PageNumberSeparator_|Optional| **Variant**|The characters (up to five) that separate individual page numbers within page references in the table of authorities. Corresponds to the \l switch for a Table of Authorities (TOA) field. If this argument is omitted, a comma and a space are used.|

### Return Value

TableOfAuthorities


## Remarks

A table of authorities is built from Table of Authorities Entry (TA) fields in a document. Use the  **MarkCitation** method to mark citations to be included in the table of authorities.


## Example

This example adds, at the beginning of the active document, a table of authorities that includes all categories.


```vb
Set myRange = ActiveDocument.Range(0, 0) 
ActiveDocument.TablesOfAuthorities.Add Range:=myRange, _ 
 Passim:= True, Category:= 0, EntrySeparator:= ", "
```


## See also


#### Concepts


[TablesOfAuthorities Collection Object](tablesofauthorities-object-word.md)

