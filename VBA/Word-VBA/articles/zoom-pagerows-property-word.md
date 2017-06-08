---
title: Zoom.PageRows Property (Word)
keywords: vbawd10.chm161873922
f1_keywords:
- vbawd10.chm161873922
ms.prod: word
api_name:
- Word.Zoom.PageRows
ms.assetid: 15db7c14-ee98-bac7-179a-018f4cb47fb9
ms.date: 06/08/2017
---


# Zoom.PageRows Property (Word)

Returns or sets the number of pages to be displayed one above the other on-screen at the same time in print layout view or print preview. Read/write  **Long** .


## Syntax

 _expression_ . **PageRows**

 _expression_ An expression that returns a **[Zoom](zoom-object-word.md)** object.


## Example

This example switches the active window to print preview and displays two pages one above the other.


```vb
PrintPreview = True 
With ActiveDocument.ActiveWindow.View.Zoom 
 .PageColumns = 1 
 .PageRows = 2 
End With
```


## See also


#### Concepts


[Zoom Object](zoom-object-word.md)

