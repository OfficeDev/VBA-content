---
title: Slide.Shapes Property (PowerPoint)
keywords: vbapp10.chm531003
f1_keywords:
- vbapp10.chm531003
ms.prod: powerpoint
api_name:
- PowerPoint.Slide.Shapes
ms.assetid: 8eaf3611-2799-835d-ecaa-c8f802256673
ms.date: 06/08/2017
---


# Slide.Shapes Property (PowerPoint)

Returns a  **[Shapes](shapes-object-powerpoint.md)** collection that represents all the elements that have been placed or inserted on the specified slide, slide master, or range of slides. Read-only.


## Syntax

 _expression_. **Shapes**

 _expression_ A variable that represents a **Slide** object.


### Return Value

Shapes


## Remarks

The  **Shapes** collection returned can contain the drawings, shapes, OLE objects, pictures, text objects, titles, headers, footers, slide numbers, and date and time objects on a slide, or on the slide image on a notes page.


## Example

This example adds a rectangle that's 100 points wide and 50 points high, and whose upper-left corner is 5 points from the left edge of slide one in the active presentation and 25 points from the top of the slide.


```vb
Set firstSlide = ActivePresentation.Slides(1)

firstSlide.Shapes.AddShape msoShapeRectangle, 5, 25, 100, 50
```

This example sets the fill texture for shape three on slide one in the active presentation.




```vb
Set newRect = ActivePresentation.Slides(1).Shapes(3)

newRect.Fill.PresetTextured msoTextureOak
```

Assuming that slide one in the active presentation contains a title, both the second and third lines of code in the following example set the title text on slide one in the presentation.




```vb
Set firstSl = ActivePresentation.Slides(1)

firstSl.Shapes.Title.TextFrame.TextRange.Text = "Some title text"

firstSl.Shapes(1).TextFrame.TextRange.Text = "Other title text"
```

Assuming that shape two on slide two in the active presentation contains a text frame, the following example adds a series of paragraphs to the slide. Note that  `Chr(13)` is used to insert paragraph marks within the text.




```vb
Set tShape = ActivePresentation.Slides(2).Shapes(2)

tShape.TextFrame.TextRange.Text = "First Item" &; Chr(13) &; _
    "Second Item" &; Chr(13) &; "Third Item"
```

For most slide layouts, the first shapes on the slide are text placeholders, and the following example accomplishes the same task as the preceding example.




```vb
Set testShape = ActivePresentation.Slides(2).Shapes.Placeholders(2)

testShape.TextFrame.TextRange.Text = "First Item" &; _
    Chr(13) &; "Second Item" &; Chr(13) &; "Third Item"
```


## See also


#### Concepts


[Slide Object](slide-object-powerpoint.md)

