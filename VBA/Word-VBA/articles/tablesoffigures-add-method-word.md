---
title: TablesOfFigures.Add Method (Word)
keywords: vbawd10.chm153223612
f1_keywords:
- vbawd10.chm153223612
ms.prod: word
api_name:
- Word.TablesOfFigures.Add
ms.assetid: 9ebee370-deeb-24b1-0fa1-a98db85e1972
ms.date: 06/08/2017
---


# TablesOfFigures.Add Method (Word)

Returns a  **TableOfFigures** object that represents a table of figures added to a document.


## Syntax

 _expression_ . **Add**( **_Range_** , **_Caption_** , **_IncludeLabel_** , **_UseHeadingStyles_** , **_UpperHeadingLevel_** , **_LowerHeadingLevel_** , **_UseFields_** , **_TableID_** , **_RightAlignPageNumbers_** , **_IncludePageNumbers_** , **_AddedStyles_** , **_UseHyperlinks_** , **_HidePageNumbersInWeb_** )

 _expression_ Required. A variable that represents a **[TablesOfFigures](tablesoffigures-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **Range object**|The range where you want the table of figures to appear.|
| _Caption_|Optional| **Variant**|The label that identifies the items you want to include in the table of figures. Corresponds to the \c switch for a Table of Contents (TOC) field. The default value is "Figure."|
| _IncludeLabel_|Optional| **Variant**| **True** to include the caption label and caption number in the table of figures. The default value is **True** .|
| _UseHeadingStyles_|Optional| **Variant**| **True** to use built-in heading styles to create the table of figures. The default value is **False** .|
| _UpperHeadingLevel_|Optional| **Variant**|The starting heading level for the table of figures, if UseHeadingStyles is set to  **True** . Corresponds to the starting value used with the \o switch for a Table of Contents (TOC) field. The default value is 1.|
| _LowerHeadingLevel_|Optional| **Variant**|The ending heading level for the table of figures, if UseHeadingStyles is set to  **True** . Corresponds to the ending value used with the \o switch for a Table of Contents (TOC) field. The default value is 9.|
| _UseFields_|Optional| **Variant**| **True** to use Table of Contents Entry (TC) fields to create the table of figures. Use the **MarkEntry** method to mark entries you want to include in the table of figures. The default value is **False** .|
| _TableID_|Optional| **Variant**|A one-letter identifier that's used to build a table of figures from Table of Contents Entry (TC) fields. Corresponds to the \f switch for a Table of Contents (TOC) field. For example, "i" builds a table of figures for an illustration.|
| _RightAlignPageNumbers_|Optional| **Variant**| **True** align page numbers with the right margin in the table of figures. The default value is **True** .|
| _IncludePageNumbers_|Optional| **Variant**| **True** if page numbers are included in the table of figures. The default value is **True** .|
| _AddedStyles_|Optional| **Variant**|The string name for additional styles used to compile the table of figures (styles other than the Heading 1 ? Heading 9 styles).|
| _UseHyperlinks_|Optional| **Variant**| **True** if entries in a table of figures should be formatted as hyperlinks when publishing to the Web. The default value is **True** .|
| _HidePageNumbersInWeb_|Optional| **Variant**| **True** if page numbers in a table of figures should be hidden when publishing to the Web. The default value is **True** .|

### Return Value

TableOfFigures


## Example

This example inserts a table of figures in the active document.


```vb
ActiveDocument.TablesOfFigures.Add Range:=Selection.Range
```


## See also


#### Concepts


[TablesOfFigures Collection Object](tablesoffigures-object-word.md)

