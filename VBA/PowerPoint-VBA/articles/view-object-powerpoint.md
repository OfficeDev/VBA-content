---
title: View Object (PowerPoint)
keywords: vbapp10.chm512000
f1_keywords:
- vbapp10.chm512000
ms.prod: powerpoint
api_name:
- PowerPoint.View
ms.assetid: 333e8b59-398d-4575-d37b-bfb1d3503089
ms.date: 06/08/2017
---


# View Object (PowerPoint)

Represents the current editing view in the specified document window.


## Remarks




 **Note**  The  **View** object can represent any of the document window views: normal view, slide view, outline view, slide sorter view, notes page view, slide master view, handout master view, or notes master view. Some properties and methods of the **View** object work only in certain views. If you try to use a property or method that's inappropriate for a **View** object, an error occurs.


## Example

Use the [View](documentwindow-view-property-powerpoint.md)property of the  **[DocumentWindow](documentwindow-object-powerpoint.md)** object to return the **View** object. The following example sets the size of window one and then sets the zoom to fit the new window size.


```vb
With Windows(1)

    .Height = 200

    .Width = 250

    .View.ZoomToFit = True

End With
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

