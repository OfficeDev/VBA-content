---
title: Endnotes.ContinuationNotice Property (Word)
keywords: vbawd10.chm155254890
f1_keywords:
- vbawd10.chm155254890
ms.prod: word
api_name:
- Word.Endnotes.ContinuationNotice
ms.assetid: 3d2007df-756e-17f9-ce7c-269fa633503b
ms.date: 06/08/2017
---


# Endnotes.ContinuationNotice Property (Word)

Returns a  **Range** object that represents the endnote continuation notice. Read-only.


## Syntax

 _expression_ . **ContinuationNotice**

 _expression_ A variable that represents an **[Endnotes](endnotes-object-word.md)** collection.


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


[Endnotes Collection Object](endnotes-object-word.md)

