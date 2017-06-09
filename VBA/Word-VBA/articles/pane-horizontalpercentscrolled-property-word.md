---
title: Pane.HorizontalPercentScrolled Property (Word)
keywords: vbawd10.chm157286413
f1_keywords:
- vbawd10.chm157286413
ms.prod: word
api_name:
- Word.Pane.HorizontalPercentScrolled
ms.assetid: db5c1e50-a910-9b5e-9767-9b2e1ce8cc94
ms.date: 06/08/2017
---


# Pane.HorizontalPercentScrolled Property (Word)

Returns or sets the horizontal scroll position as a percentage of the document width. Read/write  **Long** .


## Syntax

 _expression_ . **HorizontalPercentScrolled**

 _expression_ A variable that represents a **[Pane](pane-object-word.md)** object.


## Example

This example horizontally scrolls the active pane of the window for Document1 all the way to the left.


```vb
With Windows("Document1") 
 .Activate 
 .ActivePane.HorizontalPercentScrolled = 0 
End With
```


## See also


#### Concepts


[Pane Object](pane-object-word.md)

