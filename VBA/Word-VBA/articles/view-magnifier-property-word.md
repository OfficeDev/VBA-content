---
title: View.Magnifier Property (Word)
keywords: vbawd10.chm161808391
f1_keywords:
- vbawd10.chm161808391
ms.prod: word
api_name:
- Word.View.Magnifier
ms.assetid: a195490b-a84d-78cb-f834-f154063c1021
ms.date: 06/08/2017
---


# View.Magnifier Property (Word)

 **True** if the pointer is displayed as a magnifying glass in print preview, indicating that the user can click to zoom in on a particular area of the page or zoom out to see an entire page or spread of pages. Read/write **Boolean** .


## Syntax

 _expression_ . **Magnifier**

 _expression_ An expression that returns a **[View](view-object-word.md)** object.


## Remarks

This property generates an error if the view is not print preview.


## Example

This example switches to print preview and changes the pointer to an insertion point.


```vb
PrintPreview = True 
ActiveDocument.ActiveWindow.View.Magnifier = False
```


## See also


#### Concepts


[View Object](view-object-word.md)

