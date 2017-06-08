---
title: Pane.MinimumFontSize Property (Word)
keywords: vbawd10.chm157286415
f1_keywords:
- vbawd10.chm157286415
ms.prod: word
api_name:
- Word.Pane.MinimumFontSize
ms.assetid: 45aa3c50-ac50-c3b0-f7eb-099b4559ff43
ms.date: 06/08/2017
---


# Pane.MinimumFontSize Property (Word)

Returns or sets the minimum font size (in points) displayed for the specified pane. Read/write  **Long** .


## Syntax

 _expression_ . **MinimumFontSize**

 _expression_ An expression that returns a **[Pane](pane-object-word.md)** object.


## Remarks

This property only affects the text as shown in Web layout view. The point sizes that are displayed on the  **Formatting** toolbar and used for printing aren't changed.


## Example

This example sets the active window to online view and then sets the minimum font size for the active pane to 12 points.


```vb
With ActiveDocument.ActiveWindow 
 .View.Type = wdWebView 
 .ActivePane.MinimumFontSize = 12 
End With
```

This example returns the minimum font size for the active pane.




```
Msgbox _ 
 ActiveDocument.ActiveWindow.ActivePane.MinimumFontSize
```


## See also


#### Concepts


[Pane Object](pane-object-word.md)

