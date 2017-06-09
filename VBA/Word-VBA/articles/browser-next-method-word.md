---
title: Browser.Next Method (Word)
keywords: vbawd10.chm154009701
f1_keywords:
- vbawd10.chm154009701
ms.prod: word
api_name:
- Word.Browser.Next
ms.assetid: d1ac6216-dbd9-9b74-3ac6-133a1d83c09a
ms.date: 06/08/2017
---


# Browser.Next Method (Word)

Moves the selection to the next item indicated by the browser target. Use the  **Target** property to change the browser target.


## Syntax

 _expression_ . **Next**

 _expression_ Required. A variable that represents a **[Browser](browser-object-word.md)** object.


## Example

This example moves the insertion point just before the next comment reference marker in the active document.


```vb
With Application.Browser 
 .Target = wdBrowseComment 
 .Next 
End With
```


## See also


#### Concepts


[Browser Object](browser-object-word.md)

