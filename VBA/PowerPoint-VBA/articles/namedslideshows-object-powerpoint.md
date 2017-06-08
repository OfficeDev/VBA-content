---
title: NamedSlideShows Object (PowerPoint)
keywords: vbapp10.chm515000
f1_keywords:
- vbapp10.chm515000
ms.prod: powerpoint
api_name:
- PowerPoint.NamedSlideShows
ms.assetid: 9f20ff20-a81e-f771-5ef2-44b21ecfb055
ms.date: 06/08/2017
---


# NamedSlideShows Object (PowerPoint)

A collection of all the  **[NamedSlideShow](namedslideshow-object-powerpoint.md)** objects in the presentation. Each **NamedSlideShow** object represents a custom slide show.


## Example

Use the [NamedSlideShows](slideshowsettings-namedslideshows-property-powerpoint.md)property to return the  **NamedSlideShows** collection. Use **NamedSlideShows** (index), where index is the custom slide show name or index number, to return a single **NamedSlideShow** object. The following example deletes the custom slide show named "Quick Show."


```vb
ActivePresentation.SlideShowSettings _
    .NamedSlideShows("Quick Show").Delete
```

Use the [Add](namedslideshows-add-method-powerpoint.md)method to create a new slide show and add it to the  **NamedSlideShows** collection. The following example adds to the active presentation the named slide show "Quick Show" that contains slides 2, 7, and 9. The example then runs this custom slide show.




```vb
Dim qSlides(1 To 3) As Long

With ActivePresentation

    With .Slides

        qSlides(1) = .Item(2).SlideID

        qSlides(2) = .Item(7).SlideID

        qSlides(3) = .Item(9).SlideID

    End With

    With .SlideShowSettings

        .NamedSlideShows.Add "Quick Show", qSlides

        .RangeType = ppShowNamedSlideShow

        .SlideShowName = "Quick Show"

        .Run

    End With

End With
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

