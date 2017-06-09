---
title: TextFrame.BreakForwardLink Method (Word)
keywords: vbawd10.chm162665356
f1_keywords:
- vbawd10.chm162665356
ms.prod: word
api_name:
- Word.TextFrame.BreakForwardLink
ms.assetid: e72e07bf-cea3-3351-3fa8-aae9777babf6
ms.date: 06/08/2017
---


# TextFrame.BreakForwardLink Method (Word)

Breaks the forward link for the specified text frame, if such a link exists.


## Syntax

 _expression_ . **BreakForwardLink**

 _expression_ Required. A variable that represents a **[TextFrame](textframe-object-word.md)** object.


## Remarks

Applying this method to a shape in the middle of a chain of shapes with linked text frames will break the chain, leaving two sets of linked shapes. All of the text, however, will remain in the first series of linked shapes.


## Example

This example creates a new document adds a chain of three linked text boxes to it, and then breaks the link after the second text box.


```vb
Dim shapeTextbox1 As Shape 
Dim shapeTextbox2 As Shape 
Dim shapeTextbox3 As Shape 
 
Documents.Add 
 
Set shapeTextbox1 = ActiveDocument.Shapes.AddTextbox _ 
 (Orientation:=msoTextOrientationHorizontal, _ 
 Left:=InchesToPoints(1.5), _ 
 Top:=InchesToPoints(0.5), _ 
 Width:=InchesToPoints(1), _ 
 Height:=InchesToPoints(0.5)) 
shapeTextbox1.TextFrame.TextRange = "This is some text. " _ 
 &; "This is some more text. This is even more text." 
 
Set shapeTextbox2 = ActiveDocument.Shapes.AddTextbox _ 
 (Orientation:=msoTextOrientationHorizontal, _ 
 Left:=InchesToPoints(1.5), _ 
 Top:=InchesToPoints(1.5), _ 
 Width:=InchesToPoints(1), _ 
 Height:=InchesToPoints(0.5)) 
 
Set shapeTextbox3 = ActiveDocument.Shapes.AddTextbox _ 
 (Orientation:=msoTextOrientationHorizontal, _ 
 Left:=InchesToPoints(1.5), _ 
 Top:=InchesToPoints(2.5), _ 
 Width:=InchesToPoints(1), _ 
 Height:=InchesToPoints(0.5)) 
 
shapeTextbox1.TextFrame.Next = shapeTextbox2.TextFrame 
shapeTextbox2.TextFrame.Next = shapeTextbox3.TextFrame 
MsgBox "Textboxes 1, 2, and 3 are linked." 
shapeTextbox2.TextFrame.BreakForwardLink
```


## See also


#### Concepts


[TextFrame Object](textframe-object-word.md)

