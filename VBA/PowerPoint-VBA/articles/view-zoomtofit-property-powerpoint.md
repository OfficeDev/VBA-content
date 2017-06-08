---
title: View.ZoomToFit Property (PowerPoint)
keywords: vbapp10.chm512009
f1_keywords:
- vbapp10.chm512009
ms.prod: powerpoint
api_name:
- PowerPoint.View.ZoomToFit
ms.assetid: b35e3466-c135-bc5f-40d6-0331cf642b12
ms.date: 06/08/2017
---


# View.ZoomToFit Property (PowerPoint)

Determines whether the view is zoomed to fit the dimensions of the document window every time the document window is resized. Read/write.


## Syntax

 _expression_. **ZoomToFit**

 _expression_ A variable that represents a **View** object.


### Return Value

MsoTriState


## Remarks

This property applies only to slide view, notes page view, or master view.

When the value of the  **[Zoom](view-zoom-property-powerpoint.md)** property is explicitly set, the value of the **ZoomToFit** property is automatically set to **msoFalse**.

The value of the  **ZoomToFit** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The view is not zoomed to fit the dimensions of the document window every time the document window is resized.|
|**msoTrue**| The view is zoomed to fit the dimensions of the document window every time the document window is resized.|

## Example

The following example sets the view in document window one to slide view, with the zoom automatically set to fit the dimensions of the window.


```vb
With Windows(1)

    .ViewType = ppViewSlide

    .View.ZoomToFit = msoTrue

End With
```


## See also


#### Concepts


[View Object](view-object-powerpoint.md)

