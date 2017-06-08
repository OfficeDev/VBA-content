---
title: Sequence.ConvertToAnimateInReverse Method (PowerPoint)
keywords: vbapp10.chm651011
f1_keywords:
- vbapp10.chm651011
ms.prod: powerpoint
api_name:
- PowerPoint.Sequence.ConvertToAnimateInReverse
ms.assetid: dabea9a8-1ac5-6e2a-1932-7051efb9577d
ms.date: 06/08/2017
---


# Sequence.ConvertToAnimateInReverse Method (PowerPoint)

Determines whether text will be animated in reverse order. Returns an  **[Effect](effect-object-powerpoint.md)** object representing the text animation.


## Syntax

 _expression_. **ConvertToAnimateInReverse**( **_Effect_**, **_animateInReverse_** )

 _expression_ A variable that represents a **Sequence** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Effect_|Required|**Effect**|The animation effect to which the reversal will apply.|
| _animateInReverse_|Required|**[MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx)**|Determines the text animation order.|

### Return Value

Effect


## Example

This example creates a shape with text on a slide and adds a random animation to the shape, ensuring the shape's text animates in reverse.


```vb
Sub AnimateInReverse() 
 
    Dim sldActive As Slide 
    Dim timeMain As TimeLine 
    Dim shpRect As Shape 
 
    ' Create a slide, add a rectangular shape to the slide, and 
    ' access the slide's animation timeline. 
    With ActivePresentation 
        Set sldActive = .Slides.Add(Index:=1, Layout:=ppLayoutBlank) 
        Set shpRect = sldActive.Shapes.AddShape(Type:=msoShapeRectangle, _ 
            Left:=100, Top:=100, Width:=300, Height:=150) 
        Set timeMain = sldActive.TimeLine 
    End With 
 
    shpRect.TextFrame.TextRange.Text = "This is a rectangle." 
 
    ' Add a random animation effect to the rectangle, 
    ' and animate the text in reverse. 
    With timeMain.MainSequence 
        .ConvertToAnimateInReverse _ 
            Effect:=.AddEffect(Shape:=shpRect, effectId:=msoAnimEffectRandom), _ 
            AnimateInReverse:=msoTrue 
    End With 
 
End Sub
```


## See also


#### Concepts


[Sequence Object](sequence-object-powerpoint.md)

