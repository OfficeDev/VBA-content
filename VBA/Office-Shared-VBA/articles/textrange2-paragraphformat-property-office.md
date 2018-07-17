---
title: TextRange2.ParagraphFormat Property (Office)
ms.prod: office
api_name:
- Office.TextRange2.ParagraphFormat
ms.assetid: 68818c1a-9503-4f3f-77e1-28ac6b049c3b
ms.date: 06/08/2017
---


# TextRange2.ParagraphFormat Property (Office)

Returns a  **ParagraphFormat** object that represents paragraph formatting for the specified text. Read-only.


## Syntax

 _expression_. **ParagraphFormat**

 _expression_ An expression that returns a **TextRange2** object.


### Return Value

ParagraphFormat


## Example

This example sets the line spacing before, within, and after each paragraph in shape two on slide one in the active PowerPoint presentation.


```
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


[TextRange2 Object](textrange2-object-office.md)
#### Other resources


[TextRange2 Object Members](textrange2-members-office.md)

