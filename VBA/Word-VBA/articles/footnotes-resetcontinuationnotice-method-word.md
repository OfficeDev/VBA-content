---
title: Footnotes.ResetContinuationNotice Method (Word)
keywords: vbawd10.chm155320329
f1_keywords:
- vbawd10.chm155320329
ms.prod: word
api_name:
- Word.Footnotes.ResetContinuationNotice
ms.assetid: 7a5d4a70-bd00-2b24-619d-e7a8b50bf8f9
ms.date: 06/08/2017
---


# Footnotes.ResetContinuationNotice Method (Word)

Resets the footnote or endnote continuation notice to the default notice.


## Syntax

 _expression_ . **ResetContinuationNotice**

 _expression_ Required. A variable that represents a **[Footnotes](footnotes-object-word.md)** collection.


## Remarks

The default notice is blank (no text).


## Example

This example resets the footnote continuation notice and sets the starting number for footnote reference marks to 2 in Sales.doc.


```vb
With Documents("Sales.doc").Sections(1).Range.Footnotes 
 .ResetContinuationNotice 
 .NumberingRule = wdRestartContinuous 
 .StartingNumber = 2 
End With
```


## See also


#### Concepts


[Footnotes Collection Object](footnotes-object-word.md)

