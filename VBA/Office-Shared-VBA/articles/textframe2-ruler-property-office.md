---
title: TextFrame2.Ruler Property (Office)
ms.prod: office
api_name:
- Office.TextFrame2.Ruler
ms.assetid: 3d975982-25d3-644a-102d-aa116a606d71
ms.date: 06/08/2017
---


# TextFrame2.Ruler Property (Office)

Returns a  **Ruler2** object that represents the ruler for the specified text. Read-only


## Syntax

 _expression_. **Ruler**

 _expression_ An expression that returns a **TextFrame2** object.


## Example

The following code shows how to set a left-aligned tab stop at 2 inches (144 points) for the text in shape one on slide one in the active presentation.


```
Dim pptSlide As Slide 
Set pptSlide = ActivePresentation.Slides(1) 
pptSlide.Shapes(1).TextFrame2.Ruler.TabStops.Add ppTabStopLeft, 144 

```


## See also


#### Concepts


[TextFrame2 Object](textframe2-object-office.md)
#### Other resources


[TextFrame2 Object Members](textframe2-members-office.md)

