---
title: Range.Editors Property (Word)
keywords: vbawd10.chm157155671
f1_keywords:
- vbawd10.chm157155671
ms.prod: word
api_name:
- Word.Range.Editors
ms.assetid: fe491d3f-e559-aa3d-53ce-bf4aea0de5f8
ms.date: 06/08/2017
---


# Range.Editors Property (Word)

Returns an  **Editors** object that represents all the users authorized to modify a selection or range within a document.


## Syntax

 _expression_ . **Editors**

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


## Example

The following example gives the current user editing permission to modify the active selection.


```vb
Dim objEditor As Editor 
 
Set objEditor = Selection.Editors.Add(wdEditorCurrent)
```


## See also


#### Concepts


[Range Object](range-object-word.md)

