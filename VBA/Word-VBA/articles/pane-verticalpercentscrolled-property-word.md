---
title: Pane.VerticalPercentScrolled Property (Word)
keywords: vbawd10.chm157286414
f1_keywords:
- vbawd10.chm157286414
ms.prod: word
api_name:
- Word.Pane.VerticalPercentScrolled
ms.assetid: 1e63b432-cef1-7a3f-acef-db0d2f6221db
ms.date: 06/08/2017
---


# Pane.VerticalPercentScrolled Property (Word)

Returns or sets the vertical scroll position as a percentage of the document length. Read/write  **Long** .


## Syntax

 _expression_ . **VerticalPercentScrolled**

 _expression_ Required. A variable that represents a **[Pane](pane-object-word.md)** object.


## Example

This example vertically scrolls the active pane of the window for Document1 to the end.


```vb
With Windows("Document1") 
 .Activate 
 .ActivePane.VerticalPercentScrolled = 100 
End With
```


## See also


#### Concepts


[Pane Object](pane-object-word.md)

