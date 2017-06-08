---
title: Master.Shapes Property (PowerPoint)
keywords: vbapp10.chm533003
f1_keywords:
- vbapp10.chm533003
ms.prod: powerpoint
api_name:
- PowerPoint.Master.Shapes
ms.assetid: a4620f02-d3d2-da87-6bbc-430557365c2d
ms.date: 06/08/2017
---


# Master.Shapes Property (PowerPoint)

Returns a  **[Shapes](shapes-object-powerpoint.md)** collection that represents all the elements that have been placed or inserted on the specified slide, slide master, or range of slides. Read-only.


## Syntax

 _expression_. **Shapes**

 _expression_ A variable that represents a **Master** object.


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


[Master Object](master-object-powerpoint.md)

