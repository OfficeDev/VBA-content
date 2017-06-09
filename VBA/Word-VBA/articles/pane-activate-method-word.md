---
title: Pane.Activate Method (Word)
keywords: vbawd10.chm157286500
f1_keywords:
- vbawd10.chm157286500
ms.prod: word
api_name:
- Word.Pane.Activate
ms.assetid: 48bc8f8f-3dcb-15d4-0ab6-a83e984edbb1
ms.date: 06/08/2017
---


# Pane.Activate Method (Word)

Activates the specified pane.


## Syntax

 _expression_ . **Activate**

 _expression_ Required. A variable that represents a **[Pane](pane-object-word.md)** object.


## Example

This example splits the active window and then activates the first pane.


```vb
Sub SplitWindow() 
 With ActiveDocument.ActiveWindow 
 .SplitVertical = 50 
 .Panes(1).Activate 
 End With 
End Sub
```


## See also


#### Concepts


[Pane Object](pane-object-word.md)

