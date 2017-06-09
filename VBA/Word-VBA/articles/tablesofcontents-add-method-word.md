---
title: TablesOfContents.Add Method (Word)
keywords: vbawd10.chm152305767
f1_keywords:
- vbawd10.chm152305767
ms.prod: word
api_name:
- Word.TablesOfContents.Add
ms.assetid: a5954a99-ee43-3b8d-4d81-d12f889021b4
ms.date: 06/08/2017
---


# TablesOfContents.Add Method (Word)

Returns a  **TableOfContents** object that represents a table of contents added to a document.


## Syntax

 _expression_ . **Add**( **_Range_** , **_UseHeadingStyles_** , **_UpperHeadingLevel_** , **_LowerHeadingLevel_** , **_UseFields_** , **_TableID_** , **_RightAlignPageNumbers_** , **_IncludePageNumbers_** , **_AddedStyles_** , **_UseHyperlinks_** , **_HidePageNumbersInWeb_** , **_UseOutlineLevels_** )

 _expression_ Required. A variable that represents a **[TablesOfContents](tablesofcontents-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **Range object**|The range where you want the table of contents to appear. The table of contents replaces the range, if the range isn't collapsed.|
| _UseHeadingStyles_|Optional| **Variant**| **True** to use built-in heading styles to create the table of contents. The default value is **True** .|
| _UpperHeadingLevel_|Optional| **Variant**|The starting heading level for the table of contents. Corresponds to the starting value used with the \o switch for a Table of Contents (TOC) field. The default value is 1.|
| _LowerHeadingLevel_|Optional| **Variant**|The ending heading level for the table of contents. Corresponds to the ending value used with the \o switch for a Table of Contents (TOC) field. The default value is 9.|
| _UseFields_|Optional| **Variant**| **True** if Table of Contents Entry (TC) fields are used to create the table of contents. Use the **MarkEntry** method to mark entries to be included in the table of contents. The default value is **False** .|
| _TableID_|Optional| **Variant**|A one-letter identifier that's used to build a table of contents from TC fields. Corresponds to the \f switch for a Table of Contents (TOC) field. For example, "T" builds a table of contents from TC fields using the table identifier T. If this argument is omitted, TC fields aren't used.|
| _RightAlignPageNumbers_|Optional| **Variant**| **True** if page numbers in the table of contents are aligned with the right margin. The default value is **True** .|
| _IncludePageNumbers_|Optional| **Variant**| **True** to include page numbers in the table of contents. The default value is **True** .|
| _AddedStyles_|Optional| **Variant**|The string name for additional styles used to compile the table of contents (styles other than the Heading 1 ? Heading 9 styles). Use the  **Add** method of a **[HeadingStyles](headingstyles-object-word.md)** object to create new heading styles.|
| _UseHyperlinks_|Optional| **Variant**| **True** if entries in a table of contents should be formatted as hyperlinks when the document is being publishing to the Web. The default value is **True** .|
| _HidePageNumbersInWeb_|Optional| **Variant**| **True** if page numbers in a table of contents should be hidden when the document is being publishing to the Web. The default value is **True** .|
| _UseOutlineLevels_|Optional| **Variant**| **True** to use outline levels to create the table of contents. The default is **False** .|

### Return Value

TableOfContents


## Example

This example adds a table of contents at the beginning of the active document. The table of contents is built from paragraphs styled with the Heading 1, Heading 2, and Heading 3 styles or the custom styles myStyle and yourStyle.


```vb
Set myRange = ActiveDocument.Range(0, 0) 
ActiveDocument.TablesOfContents.Add _ 
 Range:=myRange, _ 
 UseFields:=False, _ 
 UseHeadingStyles:=True, _ 
 LowerHeadingLevel:=3, _ 
 UpperHeadingLevel:=1, _ 
 AddedStyles:="myStyle, yourStyle"
```


## See also


#### Concepts


[TablesOfContents Collection Object](tablesofcontents-object-word.md)

