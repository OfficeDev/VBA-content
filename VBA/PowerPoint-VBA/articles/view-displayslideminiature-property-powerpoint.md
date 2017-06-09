---
title: View.DisplaySlideMiniature Property (PowerPoint)
keywords: vbapp10.chm512008
f1_keywords:
- vbapp10.chm512008
ms.prod: powerpoint
api_name:
- PowerPoint.View.DisplaySlideMiniature
ms.assetid: 50781703-1e04-0e95-80d9-2b518130f3eb
ms.date: 06/08/2017
---


# View.DisplaySlideMiniature Property (PowerPoint)

Determines if and when the slide miniature window is displayed automatically. Read/write.


## Syntax

 _expression_. **DisplaySlideMiniature**

 _expression_ A variable that represents a **View** object.


### Return Value

MsoTriState


## Remarks

This property is not available in slide show view and slide sorter view. The slide miniature window isn't a member of either the  **Windows** collection or the **SlideShowWindows** collection.

The fit percentage is determined by a combination of the size of the slide pane and the size of the presentation window. To determine the fit percentage, set the  **[ZoomToFit](view-zoomtofit-property-powerpoint.md)** property to **True** and then return the value of the **[Zoom](slideshowview-zoom-property-powerpoint.md)** property.

The value of the  **DisplaySlideMiniature** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The slide miniature window is not displayed automatically.|
|**msoTrue**| The slide miniature window is displayed automatically when the document window is in black-and-white view, the slide pane is zoomed to greater than 150% of the fit percentage, or a master view is visible.|

## Example

If document window one is in slide view, this example displays the slide miniature window.


```vb
With Windows(1).View

    If .Type = ppViewSlide Then .DisplaySlideMiniature = msoTrue

End With
```


## See also


#### Concepts


[View Object](view-object-powerpoint.md)

