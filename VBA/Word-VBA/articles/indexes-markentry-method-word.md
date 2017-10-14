---
title: Indexes.MarkEntry Method (Word)
keywords: vbawd10.chm159121509
f1_keywords:
- vbawd10.chm159121509
ms.prod: word
api_name:
- Word.Indexes.MarkEntry
ms.assetid: e0c88e2b-6a5c-0ae9-3639-393a454c546b
ms.date: 06/08/2017
---


# Indexes.MarkEntry Method (Word)

Inserts an XE (Index Entry) field after the specified range. The method returns a  **Field** object representing the XE field.


## Syntax

 _expression_ . **MarkEntry**( **_Range_** , **_Entry_** , **_EntryAutoText_** , **_CrossReference_** , **_CrossReferenceAutoText_** , **_BookmarkName_** , **_Bold_** , **_Italic_** , **_Reading_** )

 _expression_ Required. A variable that represents an **[Indexes](indexes-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **Range object**|The location of the entry. The XE field is inserted after Range.|
| _Entry_|Optional| **Variant**|The text that appears in the index. To indicate a subentry, include the main entry text and the subentry text, separated by a colon (:) (for example, "Introduction:The Product").|
| _EntryAutoText_|Optional| **Variant**|The AutoText entry name that includes text for the index, table of figures, or table of contents (Entry is ignored).|
| _CrossReference_|Optional| **Variant**|A cross-reference that will appear in the index (for example, "See Apples").|
| _CrossReferenceAutoText_|Optional| **Variant**|The AutoText entry name that contains the text for a cross-reference (CrossReference is ignored).|
| _BookmarkName_|Optional| **Variant**|The name of the bookmark that marks the range of pages you want to appear in the index. If this argument is omitted, the number of the page containing the XE field appears in the index.|
| _Bold_|Optional| **Variant**| **True** to add bold formatting to the entry page numbers in the index.|
| _Italic_|Optional| **Variant**| **True** to add italic formatting to the entry page numbers in the index.|
| _Reading_|Optional| **Variant**| **True** shows an index entry in the right location when indexes are sorted phonetically (East Asian languages only).|

### Return Value

Field


## Example

This example inserts an index entry after the selection in the active document. The subentry text is the text from the selection.


```vb
If Selection.Type = wdSelectionNormal Then 
 ActiveDocument.Indexes.MarkEntry Range:=Selection.Range, _ 
 Entry:="Introduction:" &; Selection.Range.Text, Italic:=True 
End If
```


## See also


#### Concepts


[Indexes Collection Object](indexes-object-word.md)

