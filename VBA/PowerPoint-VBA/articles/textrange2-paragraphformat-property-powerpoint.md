---
title: TextRange2.ParagraphFormat Property (PowerPoint)
ms.assetid: a7f3f37e-75a2-45a9-bf97-85f8e560192c
ms.date: 06/08/2017
ms.prod: powerpoint
---


# TextRange2.ParagraphFormat Property (PowerPoint)

Returns a  **ParagraphFormat** object that represents paragraph formatting for the specified text. Read-only.


## Syntax

 _expression_. **ParagraphFormat**

 _expression_ An expression that returns a **TextRange2** object.


### Return Value

ParagraphFormat


## Example

This example sets the line spacing before, within, and after each paragraph in shape two on slide one in the active PowerPoint presentation.


```vb
With Application.ActivePresentation.Slides(2).Shapes(2) 
 With .TextFrame.TextRange2.ParagraphFormat 
 .LineRuleWithin = msoTrue 
 .SpaceWithin = 1.4 
 .LineRuleBefore = msoTrue 
 .SpaceBefore = 0.25 
 .LineRuleAfter = msoTrue 
 .SpaceAfter = 0.75 
 End With 
End With
```


## See also


#### Concepts


[TextRange2 Object (PowerPoint)](textrange2-object-powerpoint.md)


