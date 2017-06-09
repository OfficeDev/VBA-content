---
title: View.Paste Method (PowerPoint)
keywords: vbapp10.chm512005
f1_keywords:
- vbapp10.chm512005
ms.prod: powerpoint
api_name:
- PowerPoint.View.Paste
ms.assetid: e7878c74-92d7-8993-9b46-8647c1b59b15
ms.date: 06/08/2017
---


# View.Paste Method (PowerPoint)

Pastes the contents of the Clipboard into the active view. Attempting to paste an object into a view that won't accept it causes an error. 


## Syntax

 _expression_. **Paste**

 _expression_ A variable that represents a **View** object.


## Remarks

Attempting to paste an object into a view that won't accept it causes an error. 

Use the  **[ViewType](documentwindow-viewtype-property-powerpoint.md)** property to set the view for a window before pasting the Clipboard contents into it. The following table shows what you can paste into each view.



|**Into this view**|**You can paste the following from the Clipboard**|
|:-----|:-----|
|Slide view or notes page view|Shapes, text, or entire slides. If you paste a slide from the Clipboard, an image of the slide will be inserted onto the slide, master, or notes page as an embedded object. If one shape is selected, the pasted text will be appended to the shape's text; if text is selected, the pasted text will replace the selection; if anything else is selected, the pasted text will be placed in it is own text frame. Pasted shapes will be added to the top of the z-order and won't replace selected shapes.|
|Outline view|Text or entire slides. You cannot paste shapes into outline view. A pasted slide will be inserted before the slide that contains the cursor.|
|Slide sorter view|Entire slides. You cannot paste shapes or text into slide sorter view. A pasted slide will be inserted at the cursor or after the last slide selected in the presentation.|

## Example

This example copies the selection in window one to the Clipboard and copies it into the view in window two. If the Clipboard contents cannot be pasted into the view in window two ? for example, if you try to paste a shape into slide sorter view ? this example fails.


```
Windows(1).Selection.Copy

Windows(2).View.Paste
```

This example copies the selection in window one to the Clipboard, makes sure that window one is in slide view, and then copies the Clipboard contents into the view in window two.




```
Windows(1).Selection.Copy

With Windows(2)

    .ViewType = ppViewSlide

    .View.Paste

End With


```


## See also


#### Concepts


[View Object](view-object-powerpoint.md)

