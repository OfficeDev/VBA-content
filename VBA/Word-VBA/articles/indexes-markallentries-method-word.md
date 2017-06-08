---
title: Indexes.MarkAllEntries Method (Word)
keywords: vbawd10.chm159121510
f1_keywords:
- vbawd10.chm159121510
ms.prod: word
api_name:
- Word.Indexes.MarkAllEntries
ms.assetid: bd2fb9b9-7a10-6f35-b691-f8c11542a711
ms.date: 06/08/2017
---


# Indexes.MarkAllEntries Method (Word)

Inserts an XE (Index Entry) field after all instances of the text in  **Range** .


## Syntax

 _expression_ . **MarkAllEntries**( **_Range_** , **_Entry_** , **_EntryAutoText_** , **_CrossReference_** , **_CrossReferenceAutoText_** , **_BookmarkName_** , **_Bold_** , **_Italic_** )

 _expression_ Required. A variable that represents an **[Indexes](indexes-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **Range object**|The range whose text is marked with an XE field throughout the document.|
| _Entry_|Optional| **Variant**|The text you want to appear in the index, in the form MainEntry[:Subentry].|
| _EntryAutoText_|Optional| **Variant**|The AutoText entry that contains the text you want to appear in the index (if this argument is specified, Entry is ignored).|
| _CrossReference_|Optional| **Variant**|A cross-reference that will appear in the index.|
| _CrossReferenceAutoText_|Optional| **Variant**|The name of the AutoText entry that contains the text for a cross-reference (if this argument is specified, CrossReference is ignored).|
| _BookmarkName_|Optional| **Variant**|The bookmark name that marks the range of pages you want to appear in the index. If this argument is omitted, the number of the page that contains the XE field appears in the index.|
| _Bold_|Optional| **Variant**| **True** to add bold formatting to page numbers for index entries.|
| _Italic_|Optional| **Variant**| **True** to add italic formatting to page numbers for index entries.|

## Example

This example marks the selected text with TA fields throughout the active document and then updates the first index in the document. The entry text in the index matches the selected text.


```vb
If Selection.Type = wdSelectionNormal Then 
 ActiveDocument.Indexes.MarkAllEntries _ 
 Range:=Selection.Range, _ 
 Entry:=Selection.Range.Text, Italic:=True 
 ActiveDocument.Indexes(1).Update 
End If
```


## See also


#### Concepts


[Indexes Collection Object](indexes-object-word.md)

