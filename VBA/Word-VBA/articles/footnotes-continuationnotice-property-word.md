---
title: Footnotes.ContinuationNotice Property (Word)
keywords: vbawd10.chm155320426
f1_keywords:
- vbawd10.chm155320426
ms.prod: word
api_name:
- Word.Footnotes.ContinuationNotice
ms.assetid: 355a8bc1-3cf6-51e7-27f6-f3ff2b708fca
ms.date: 06/08/2017
---


# Footnotes.ContinuationNotice Property (Word)

Returns a  **Range** object that represents the footnote continuation notice. Read-only.


## Syntax

 _expression_ . **ContinuationNotice**

 _expression_ A variable that represents a **[Footnotes](footnotes-object-word.md)** collection.


## Example

This example replaces the current footnote continuation notice with the text "Continued...".


```vb
With ActiveDocument.Footnotes.ContinuationNotice 
 .Delete 
 .InsertBefore "Continued..." 
End With
```


## See also


#### Concepts


[Footnotes Collection Object](footnotes-object-word.md)

