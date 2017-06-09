---
title: Shapes.Paste Method (PowerPoint)
keywords: vbapp10.chm543026
f1_keywords:
- vbapp10.chm543026
ms.prod: powerpoint
api_name:
- PowerPoint.Shapes.Paste
ms.assetid: 8aa534f8-bd59-3945-cc1f-45ffc3883bf7
ms.date: 06/08/2017
---


# Shapes.Paste Method (PowerPoint)

Pastes the shapes, slides, or text on the Clipboard into the specified  **Shapes** collection, at the top of the z-order. Each pasted object becomes a member of the specified **Shapes** collection. If the Clipboard contains entire slides, the slides will be pasted as shapes that contain the images of the slides. If the Clipboard contains a text range, the text will be pasted into a newly created **TextFrame** shape. Returns a **[ShapeRange](shaperange-object-powerpoint.md)** object that represents the pasted objects.


## Syntax

 _expression_. **Paste**

 _expression_ A variable that represents a **Shapes** object.


### Return Value

ShapeRange


## Remarks

Use the  **[ViewType](documentwindow-viewtype-property-powerpoint.md)** property to set the view for a window before pasting the Clipboard contents into it. The following table shows what you can paste into each view.



|**Into this view**|**You can paste the following from the Clipboard**|
|:-----|:-----|
|Slide view or notes page view|Shapes, text, or entire slides. If you paste a slide from the Clipboard, an image of the slide will be inserted onto the slide, master, or notes page as an embedded object. If one shape is selected, the pasted text will be appended to the shape's text; if text is selected, the pasted text will replace the selection; if anything else is selected, the pasted text will be placed in it is own text frame. Pasted shapes will be added to the top of the z-order and won't replace selected shapes.|
|Outline view|Text or entire slides. You cannot paste shapes into outline view. A pasted slide will be inserted before the slide that contains the cursor.|
|Slide sorter view|Entire slides. You cannot paste shapes or text into slide sorter view. A pasted slide will be inserted at the cursor or after the last slide selected in the presentation.|

## Example

This example copies shape one on slide one in the active presentation to the Clipboard and then pastes it into slide two.


```vb
With ActivePresentation

    .Slides(1).Shapes(1).Copy

    .Slides(2).Shapes.Paste

End With
```

This example cuts the text in shape one on slide one in the active presentation, places it on the Clipboard, and then pastes it after the first word in shape two on the same slide.




```vb
With ActivePresentation.Slides(1)

    .Shapes(1).TextFrame.TextRange.Cut

    .Shapes(2).TextFrame.TextRange.Words(1).InsertAfter.Paste

End With
```


## See also


#### Concepts


[Shapes Object](shapes-object-powerpoint.md)

