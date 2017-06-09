---
title: TextRange2.Length Property (Office)
ms.prod: office
api_name:
- Office.TextRange2.Length
ms.assetid: 3b873f1f-5120-3832-1d34-b8c0f668bba3
ms.date: 06/08/2017
---


# TextRange2.Length Property (Office)

Get a Long that represents the length of a text range. Read-only.


## Syntax

 _expression_. **Length**

 _expression_ An expression that returns a **TextRange2** object.


### Return Value

Long


## Example

This example adds a shape with text and rotates the shape without rotating the text in the active PowerPoint presentation.


```
Sub SetTextRange() 
 Dim shpStar As Shape 
 Dim sldOne As Slide 
 Dim effNew As Effect 
 
 Set sldOne = ActivePresentation.Slides(1) 
 Set shpStar = sldOne.Shapes.AddShape(Type:=msoShape5pointStar, _ 
 Left:=32, Top:=32, Width:=300, Height:=300) 
 
 shpStar.TextFrame.TextRange2.Text = "Animated shape." 
 
 Set effNew = sldOne.TimeLine.MainSequence.AddEffect(Shape:=shpStar, _ 
 EffectId:=msoAnimEffectPath5PointStar, Level:=msoAnimateTextByAllLevels, _ 
 Trigger:=msoAnimTriggerAfterPrevious) 
 With effNew 
 If .TextRangeStart = 0 And .TextRangeLength > 0 Then 
 With .Behaviors.Add(Type:=msoAnimTypeRotation).RotationEffect 
 .From = 0 
 .To = 360 
 End With 
 .Timing.AutoReverse = msoTrue 
 End If 
 End With 
End Sub 

```


## See also


#### Concepts


[TextRange2 Object](textrange2-object-office.md)
#### Other resources


[TextRange2 Object Members](textrange2-members-office.md)

