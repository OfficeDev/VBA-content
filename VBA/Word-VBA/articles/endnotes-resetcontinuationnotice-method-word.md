---
title: Endnotes.ResetContinuationNotice Method (Word)
keywords: vbawd10.chm155254793
f1_keywords:
- vbawd10.chm155254793
ms.prod: word
api_name:
- Word.Endnotes.ResetContinuationNotice
ms.assetid: b7565c90-6aaa-1154-ce45-61b43149ecb0
ms.date: 06/08/2017
---


# Endnotes.ResetContinuationNotice Method (Word)

Resets the endnote continuation notice to the default notice.


## Syntax

 _expression_ . **ResetContinuationNotice**

 _expression_ Required. A variable that represents an **[Endnotes](endnotes-object-word.md)** collection.


## Remarks

The default notice is blank (no text).


## Example

This example resets the endnote continuation notice for the active document.


```vb
ActiveDocument.Endnotes.ResetContinuationNotice
```


## See also


#### Concepts


[Endnotes Collection Object](endnotes-object-word.md)

