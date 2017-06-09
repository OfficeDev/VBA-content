---
title: Index.NumberOfColumns Property (Word)
keywords: vbawd10.chm159186948
f1_keywords:
- vbawd10.chm159186948
ms.prod: word
api_name:
- Word.Index.NumberOfColumns
ms.assetid: e61eaa82-d7b5-84bc-dfe9-1e410d1ec6af
ms.date: 06/08/2017
---


# Index.NumberOfColumns Property (Word)

Sets or returns the number of columns for each page of an index. Read/write  **Long** .


## Syntax

 _expression_ . **NumberOfColumns**

 _expression_ An expression that an **[Index](index-object-word.md)** object.


## Remarks

Specifying 0 (zero) sets the number of columns in the index to the same number as in the document.


## Example

This example sets the number of columns in the first index to the same number as in the active document.


```vb
ActiveDocument.Indexes(1).NumberOfColumns = 0
```

This example sets a two-column format for each index in the active document.




```vb
For Each myIndex In ActiveDocument.Indexes 
 myIndex.NumberOfColumns = 2 
Next myIndex
```


## See also


#### Concepts


[Index Object](index-object-word.md)

