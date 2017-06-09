---
title: NamedSlideShow Object (PowerPoint)
keywords: vbapp10.chm516000
f1_keywords:
- vbapp10.chm516000
ms.prod: powerpoint
api_name:
- PowerPoint.NamedSlideShow
ms.assetid: 2f5ddeeb-ecf5-50da-99b9-b38e789fd5bb
ms.date: 06/08/2017
---


# NamedSlideShow Object (PowerPoint)

Represents a custom slide show, which is a named subset of slides in a presentation. 


## Remarks

The  **NamedSlideShow** object is a member of the **[NamedSlideShows](namedslideshows-object-powerpoint.md)** collection. The **NamedSlideShows** collection contains all the named slide shows in the presentation.


## Example

Use  **NamedSlideShows** (index), where index is the custom slide show name or index number, to return a single **NamedSlideShow** object. The following example deletes the custom slide show named "Quick Show."


```vb
ActivePresentation.SlideShowSettings _
    .NamedSlideShows("Quick Show").Delete
```

Use the [SlideIDs](namedslideshow-slideids-property-powerpoint.md)property to return an array that contains the unique slide IDs for all the slides in the specified custom show. The following example displays the slide IDs for the slides in the custom slide show named "Quick Show."




```
idArray = ActivePresentation.SlideShowSettings _
    .NamedSlideShows("Quick Show").SlideIDs

For i = 1 To UBound(idArray)
    MsgBox idArray(i)
Next
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

