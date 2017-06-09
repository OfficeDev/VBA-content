---
title: Subdocuments.AddFromRange Method (Word)
keywords: vbawd10.chm159907941
f1_keywords:
- vbawd10.chm159907941
ms.prod: word
api_name:
- Word.Subdocuments.AddFromRange
ms.assetid: ca205880-99d4-2cc5-cb45-3fd8fd60cf36
ms.date: 06/08/2017
---


# Subdocuments.AddFromRange Method (Word)

Creates one or more subdocuments from the text in the specified range and returns a  **SubDocument** object.


## Syntax

 _expression_ . **AddFromRange**( **_Range_** )

 _expression_ Required. A variable that represents a **[Subdocuments](subdocuments-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **Range object**|The range used to create one or more subdocuments.|

### Return Value

SubDocument


## Remarks

The Range argument must begin with one of the built-in heading level styles (for example, Heading 1). Subdocuments are created at each paragraph formatted with the same heading format used at the beginning of the range. Subdocument files are saved when the master document is saved and are automatically named using text from the first line in the file.


## Example

This example creates one or more subdocuments from the selection.


```vb
ActiveDocument.ActiveWindow.View.Type = wdMasterView 
ActiveDocument.SubDocuments.AddFromRange Range:=Selection.Range
```


## See also


#### Concepts


[Subdocuments Collection Object](subdocuments-object-word.md)

