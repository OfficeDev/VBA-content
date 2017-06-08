---
title: Selection.Editors Property (Word)
keywords: vbawd10.chm158662969
f1_keywords:
- vbawd10.chm158662969
ms.prod: word
api_name:
- Word.Selection.Editors
ms.assetid: ba750568-88c9-9ed8-61ee-45f89dfa4dea
ms.date: 06/08/2017
---


# Selection.Editors Property (Word)

Returns an  **[Editors](editors-object-word.md)** object that represents all the users authorized to modify a selection within a document.


## Syntax

 _expression_ . **Editors**

 _expression_ A variable that represents a **[Selection](selection-object-word.md)** object.


## Example

The following example gives the current user editing permission to modify the active selection.


```vb
Dim objEditor As Editor 
 
Set objEditor = Selection.Editors.Add(wdEditorCurrent)
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

