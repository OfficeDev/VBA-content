---
title: Border.Inside Property (Word)
keywords: vbawd10.chm154861570
f1_keywords:
- vbawd10.chm154861570
ms.prod: word
api_name:
- Word.Border.Inside
ms.assetid: 73a38a3c-6c24-36f2-c6c6-8b4d2f61dc07
ms.date: 06/08/2017
---


# Border.Inside Property (Word)

 **True** if an inside border can be applied to the specified object. Read-only **Boolean** .


## Syntax

 _expression_ . **Inside**

 _expression_ An expression that returns a **[Border](border-object-word.md)** object.


## Example

If the current selection supports inside borders (that is, if multiple paragraphs or cells are selected), this example applies a single inside border.


```vb
Dim borderLoop As Border 
 
For Each borderLoop In Selection.Borders 
 If borderLoop.Inside = True Then _ 
 borderLoop.LineStyle = wdLineStyleSingle 
Next borderLoop
```


## See also


#### Concepts


[Border Object](border-object-word.md)

