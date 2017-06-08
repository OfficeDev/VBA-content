---
title: Pane.Close Method (Word)
keywords: vbawd10.chm157286501
f1_keywords:
- vbawd10.chm157286501
ms.prod: word
api_name:
- Word.Pane.Close
ms.assetid: 05e27bd2-151e-a972-9da1-13dc1d81f513
ms.date: 06/08/2017
---


# Pane.Close Method (Word)

Closes the specified Mail Merge data source, pane, or task.


## Syntax

 _expression_ . **Close**

 _expression_ Required. A variable that represents a **[Pane](pane-object-word.md)** object.


## Example

This example closes the active pane if the active window is split.


```vb
If ActiveDocument.ActiveWindow.Panes.Count >= 2 Then _ 
 ActiveDocument.ActiveWindow.ActivePane.Close
```


## See also


#### Concepts


[Pane Object](pane-object-word.md)

