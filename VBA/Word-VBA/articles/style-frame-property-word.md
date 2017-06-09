---
title: Style.Frame Property (Word)
keywords: vbawd10.chm153878539
f1_keywords:
- vbawd10.chm153878539
ms.prod: word
api_name:
- Word.Style.Frame
ms.assetid: 4e6d821d-bff8-5807-f4dc-1a1c7b4150b7
ms.date: 06/08/2017
---


# Style.Frame Property (Word)

Returns a  **[Frame](frame-object-word.md)** object that represents the frame formatting for the specified style. Read-only.


## Syntax

 _expression_ . **Frame**

 _expression_ A variable that represents a **[Style](style-object-word.md)** object.


## Example

This example creates a style with frame formatting and then applies the style to the first paragraph in the selection.


```vb
Dim styleNew As Style 
 
Set styleNew = ActiveDocument.Styles _ 
 .Add(Name:="frame", Type:=wdStyleTypeParagraph) 
With styleNew.Frame 
 .RelativeHorizontalPosition = _ 
 wdRelativeHorizontalPositionMargin 
 .HeightRule = wdFrameAuto 
 .WidthRule = wdFrameAuto 
 .TextWrap = True 
End With 
Selection.Paragraphs(1).Range.Style = "frame"
```


## See also


#### Concepts


[Style Object](style-object-word.md)

