---
title: View.ShowOptionalBreaks Property (Word)
keywords: vbawd10.chm161808415
f1_keywords:
- vbawd10.chm161808415
ms.prod: word
api_name:
- Word.View.ShowOptionalBreaks
ms.assetid: e8d6d19e-9183-52cb-df79-d3678e75a461
ms.date: 06/08/2017
---


# View.ShowOptionalBreaks Property (Word)

 **True** if Microsoft Word displays optional line breaks. Read/write **Boolean** .


## Syntax

 _expression_ . **ShowOptionalBreaks**

 _expression_ An expression that returns a **[View](view-object-word.md)** object.


## Example

This example displays the optional line breaks in the active window.


```vb
ActiveDocument.ActiveWindow.View.ShowOptionalBreaks = True
```


## See also


#### Concepts


[View Object](view-object-word.md)

