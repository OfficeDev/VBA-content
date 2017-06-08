---
title: Selection.CreateTextbox Method (Word)
keywords: vbawd10.chm158663179
f1_keywords:
- vbawd10.chm158663179
ms.prod: word
api_name:
- Word.Selection.CreateTextbox
ms.assetid: e3c567ee-949f-5e87-43c2-633cdae334b0
ms.date: 06/08/2017
---


# Selection.CreateTextbox Method (Word)

Adds a default-size text box around the selection.


## Syntax

 _expression_ . **CreateTextbox**

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

If the selection is an insertion point, this method changes the pointer to a cross-hair pointer so that the user can draw a text box.

Using this method is equivalent to clicking the  **Text Box** button on the **Drawing** toolbar. A text box is a rectangle with an associated text frame.


## Example

This example adds a text box around the selection and then changes the text box's line style.


```vb
If Selection.Type = wdSelectionNormal Then 
 Selection.CreateTextbox 
 Selection.ShapeRange(1).Line.DashStyle =msoLineDashDot 
End If
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

