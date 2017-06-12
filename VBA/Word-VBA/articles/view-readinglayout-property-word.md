---
title: View.ReadingLayout Property (Word)
keywords: vbawd10.chm161808429
f1_keywords:
- vbawd10.chm161808429
ms.prod: word
api_name:
- Word.View.ReadingLayout
ms.assetid: e53d6913-0c2c-2933-384a-31b1e8ab4986
ms.date: 06/08/2017
---


# View.ReadingLayout Property (Word)

Sets or returns a  **Boolean** that represents whether a document is being viewed in reading layout view. .


## Syntax

 _expression_ . **ReadingLayout**

 _expression_ An expression that returns a **[View](view-object-word.md)** object.


## Remarks

 **True** switches the document to reading layout view. **False** closes reading layout view.


## Example

The following example closes reading layout view.


```vb
ActiveDocument.ActiveWindow.View.ReadingLayout = False
```


## See also


#### Concepts


[View Object](view-object-word.md)

