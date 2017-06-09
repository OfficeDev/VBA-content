---
title: TextFrame.Orientation Property (Word)
keywords: vbawd10.chm162660456
f1_keywords:
- vbawd10.chm162660456
ms.prod: word
api_name:
- Word.TextFrame.Orientation
ms.assetid: 480b0ebd-c39c-0159-06a1-c909111d9486
ms.date: 06/08/2017
---


# TextFrame.Orientation Property (Word)

Returns or sets the orientation of the text inside the frame. Read/write  **MsoTextOrientation** .


## Syntax

 _expression_ . **Orientation**

 _expression_ Required. A variable that represents a **[TextFrame](textframe-object-word.md)** object.


## Remarks

Some of the  **MsoTextOrientation** constants may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.

You can set the orientation for a text frame or for a range or selection that happens to occur inside a text frame. For information about the difference between a text frame and a text box, see the  **TextFrame** object.


## Example

This example creates a new document, inserts text into it, uses this text to create a text box, and then sets the orientation of the text frame so that the text slopes upward.


```vb
Set mydoc = Documents.Add 
Selection.TypeText "This is some text." 
mydoc.Content.Select 
Selection.CreateTextbox 
mydoc.Shapes(1).TextFrame.Orientation = msoTextOrientationUpward
```


## See also


#### Concepts


[TextFrame Object](textframe-object-word.md)

