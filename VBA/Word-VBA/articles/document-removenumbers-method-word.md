---
title: Document.RemoveNumbers Method (Word)
keywords: vbawd10.chm158007436
f1_keywords:
- vbawd10.chm158007436
ms.prod: word
api_name:
- Word.Document.RemoveNumbers
ms.assetid: 2f481145-f1ef-7b80-0287-3c14a5f3d2d5
ms.date: 06/08/2017
---


# Document.RemoveNumbers Method (Word)

Removes numbers or bullets from the specified document.


## Syntax

 _expression_ . **RemoveNumbers**( **_NumberType_** )

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _NumberType_|Optional| **WdNumberType**|The type of number to be removed.|

## Example

This example removes the numbers from the beginning of any numbered paragraphs in the active document.


```vb
ActiveDocument.RemoveNumbers wdNumberParagraph
```

This example removes the bullets or numbers from the third list in MyDocument.doc.




```vb
If Documents("MyDocument.doc").Lists.Count >= 3 Then 
 Documents("MyDocument.doc").Lists(3).RemoveNumbers 
End If
```


## See also


#### Concepts


[Document Object](document-object-word.md)

