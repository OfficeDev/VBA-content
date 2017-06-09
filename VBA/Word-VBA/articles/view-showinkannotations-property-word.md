---
title: View.ShowInkAnnotations Property (Word)
keywords: vbawd10.chm161808432
f1_keywords:
- vbawd10.chm161808432
ms.prod: word
api_name:
- Word.View.ShowInkAnnotations
ms.assetid: 5e022729-7e0e-4843-adbe-cd774c2d8e23
ms.date: 06/08/2017
---


# View.ShowInkAnnotations Property (Word)

Returns or sets  **Boolean** that shows or hides handwritten ink annotations. **True** displays ink annotations. **False** hides ink annotations.


## Syntax

 _expression_ . **ShowInkAnnotations**

 _expression_ A variable that represents a **[View](view-object-word.md)** object.


## Remarks

To work with ink annotations, you must be running Microsoft Word on a tablet computer. For more information on adding handwritten ink annotations to a document, see "Mark up a document with ink annotations" in Microsoft Word Help.


## Example

The following example shows all handwritten ink annotations in the active document.


```vb
ActiveDocument.ActiveWindow.View.ShowInkAnnotations = True
```


## See also


#### Concepts


[View Object](view-object-word.md)

