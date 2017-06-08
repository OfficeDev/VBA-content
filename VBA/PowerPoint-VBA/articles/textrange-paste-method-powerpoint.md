---
title: TextRange.Paste Method (PowerPoint)
keywords: vbapp10.chm569030
f1_keywords:
- vbapp10.chm569030
ms.prod: powerpoint
api_name:
- PowerPoint.TextRange.Paste
ms.assetid: 4bbaa1f3-206e-2009-11f0-5abde24517c6
ms.date: 06/08/2017
---


# TextRange.Paste Method (PowerPoint)

Pastes the text on the Clipboard into the specified text range, and returns a  **TextRange** object that represents the pasted text.


## Syntax

 _expression_. **Paste**

 _expression_ A variable that represents a **TextRange** object.


### Return Value

TextRange


## Remarks

Use the  **[ViewType](documentwindow-viewtype-property-powerpoint.md)** property to set the view for a window before pasting the Clipboard contents into it. The following table shows what you can paste into each view.



|**Into this view**|**You can paste the following from the Clipboard**|
|:-----|:-----|
|Slide view or notes page view|Shapes, text, or entire slides. If you paste a slide from the Clipboard, an image of the slide will be inserted onto the slide, master, or notes page as an embedded object. If one shape is selected, the pasted text will be appended to the shape's text; if text is selected, the pasted text will replace the selection; if anything else is selected, the pasted text will be placed in it is own text frame. Pasted shapes will be added to the top of the z-order and won't replace selected shapes.|
|Outline view|Text or entire slides. You cannot paste shapes into outline view. A pasted slide will be inserted before the slide that contains the cursor.|
|Slide sorter view|Entire slides. You cannot paste shapes or text into slide sorter view. A pasted slide will be inserted at the cursor or after the last slide selected in the presentation.|

## See also


#### Concepts


[TextRange Object](textrange-object-powerpoint.md)

