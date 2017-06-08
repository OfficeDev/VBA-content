---
title: TextFrame.BreakForwardLink Method (Publisher)
keywords: vbapb10.chm3866661
f1_keywords:
- vbapb10.chm3866661
ms.prod: publisher
api_name:
- Publisher.TextFrame.BreakForwardLink
ms.assetid: 60a7a798-ebd3-e00d-032d-685dd0d5a042
ms.date: 06/08/2017
---


# TextFrame.BreakForwardLink Method (Publisher)

Breaks the forward link for the specified text frame, if such a link exists.


## Syntax

 _expression_. **BreakForwardLink**

 _expression_A variable that represents a  **TextFrame** object.


## Remarks

Applying this method to a shape in the middle of a chain of shapes with linked text frames will break the chain, leaving two sets of linked shapes. All of the text, however, will remain in the first series of linked shapes.


## Example

This example creates a new publication, adds a chain of three linked text boxes to it, and then breaks the link after the second text box.


```vb
Sub BreakTextLink() 
 Dim shpTextbox1 As Shape 
 Dim shpTextbox2 As Shape 
 Dim shpTextbox3 As Shape 
 
 Set shpTextbox1 = ActiveDocument.Pages(1).Shapes.AddTextbox _ 
 (Orientation:=msoTextOrientationHorizontal, _ 
 Left:=72, Top:=36, Width:=72, Height:=36) 
 shpTextbox1.TextFrame.TextRange = "This is some text. " _ 
 &; "This is some more text. This is even more text. " _ 
 &; "And this is some more text and even more text." 
 
 Set shpTextbox2 = ActiveDocument.Pages(1).Shapes.AddTextbox _ 
 (Orientation:=msoTextOrientationHorizontal, _ 
 Left:=72, Top:=108, Width:=72, Height:=36) 
 
 Set shpTextbox3 = ActiveDocument.Pages(1).Shapes.AddTextbox _ 
 (Orientation:=msoTextOrientationHorizontal, _ 
 Left:=72, Top:=180, Width:=72, Height:=36) 
 
 shpTextbox1.TextFrame.NextLinkedTextFrame = shpTextbox2.TextFrame 
 shpTextbox2.TextFrame.NextLinkedTextFrame = shpTextbox3.TextFrame 
 MsgBox "Textboxes 1, 2, and 3 are linked." 
 shpTextbox2.TextFrame.BreakForwardLink 
End Sub
```


