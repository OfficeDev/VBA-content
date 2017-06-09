---
title: Window.Thumbnails Property (Word)
keywords: vbawd10.chm157417509
f1_keywords:
- vbawd10.chm157417509
ms.prod: word
api_name:
- Word.Window.Thumbnails
ms.assetid: 2979b109-e2e6-34de-539b-53c46b0d0c55
ms.date: 06/08/2017
---


# Window.Thumbnails Property (Word)

Sets or returns a  **Boolean** that represents whether thumbnail images of the pages in a document are displayed along the left side of the Microsoft Word document window.


## Syntax

 _expression_ . **Thumbnails**

 _expression_ An expression that returns a **[Window](window-object-word.md)** object.


## Example

The following example displays thumbnail images of the pages in the active document.


```vb
ActiveDocument.ActiveWindow.Thumbnails = True
```


## See also


#### Concepts


[Window Object](window-object-word.md)

