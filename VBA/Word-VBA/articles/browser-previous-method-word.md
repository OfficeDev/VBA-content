---
title: Browser.Previous Method (Word)
keywords: vbawd10.chm154009702
f1_keywords:
- vbawd10.chm154009702
ms.prod: word
api_name:
- Word.Browser.Previous
ms.assetid: b23b637e-b7ee-05ae-dd7a-9f97ca2e6d7c
ms.date: 06/08/2017
---


# Browser.Previous Method (Word)

Moves the selection to the previous item indicated by the browser target. Use the  **Target** property to change the browser target.


## Syntax

 _expression_ . **Previous**

 _expression_ Required. A variable that represents a **[Browser](browser-object-word.md)** object.


## Example

This example moves the insertion point into the first cell (the cell in the upper-left corner) of the previous table.


```vb
With Application.Browser 
 .Target = wdBrowseTable 
 .Previous 
End With
```


## See also


#### Concepts


[Browser Object](browser-object-word.md)

