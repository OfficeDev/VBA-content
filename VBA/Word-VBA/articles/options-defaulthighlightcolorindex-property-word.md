---
title: Options.DefaultHighlightColorIndex Property (Word)
keywords: vbawd10.chm162988306
f1_keywords:
- vbawd10.chm162988306
ms.prod: word
api_name:
- Word.Options.DefaultHighlightColorIndex
ms.assetid: 1171cc44-54c9-0a39-c90f-ebdebebdde26
ms.date: 06/08/2017
---


# Options.DefaultHighlightColorIndex Property (Word)

Returns or sets the color used to highlight text formatted with the  **Highlight** button ( **Formatting** toolbar). Read/write **WdColorIndex** .


## Syntax

 _expression_ . **DefaultHighlightColorIndex**

 _expression_ Required. A variable that represents an **[Options](options-object-word.md)** collection.


## Example

This example sets the default highlight color to bright green. The new color doesn't apply to any previously highlighted text.


```
Options.DefaultHighlightColorIndex = wdBrightGreen
```

This example returns the current default highlight color index.




```vb
Dim lngTemp As Long 
 
lngTemp = Options.DefaultHighlightColorIndex
```


## See also


#### Concepts


[Options Object](options-object-word.md)

