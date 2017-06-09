---
title: Document.ClosePrintPreview Method (Word)
keywords: vbawd10.chm158007554
f1_keywords:
- vbawd10.chm158007554
ms.prod: word
api_name:
- Word.Document.ClosePrintPreview
ms.assetid: 8b4beae3-1893-5dbf-4463-bbce0c63b8ee
ms.date: 06/08/2017
---


# Document.ClosePrintPreview Method (Word)

Switches the specified document from print preview to the previous view.


## Syntax

 _expression_ . **ClosePrintPreview**

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

If the specified document isn't in print preview, an error occurs.


## Example

This example switches the active window from print preview to normal view.


```vb
If ActiveDocument.PrintPreview = True Then _ 
 ActiveDocument.ClosePrintPreview 
ActiveDocument.ActiveWindow.View.Type = wdNormalView
```


## See also


#### Concepts


[Document Object](document-object-word.md)

