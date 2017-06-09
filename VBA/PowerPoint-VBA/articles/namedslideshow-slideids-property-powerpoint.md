---
title: NamedSlideShow.SlideIDs Property (PowerPoint)
keywords: vbapp10.chm516005
f1_keywords:
- vbapp10.chm516005
ms.prod: powerpoint
api_name:
- PowerPoint.NamedSlideShow.SlideIDs
ms.assetid: 69c2a31e-bfb1-1a00-777f-4f5c46023ba0
ms.date: 06/08/2017
---


# NamedSlideShow.SlideIDs Property (PowerPoint)

Returns an array of slide IDs for the specified named slide show. Read-only.


## Syntax

 _expression_. **SlideIDs**

 _expression_ A variable that represents a **NamedSlideShow** object.


### Return Value

Variant


## Example

This example adds the current slide in the active window to the custom slide show named "Marketing Short Version." Note that to save a modified version of the custom slide show, you must delete the original custom show and then add it again, using the same name. Also note that if you want to resize an array contained in a  **Variant** variable, you must explicitly declare the variable before attempting to resize its array.


```vb
'NOTE - The following code line is NOT optional.
'Can't redim array without this
Dim customShowSlideIDs As Variant
Dim customShowToExpand As NamedSlideShow

customShowName = "Marketing Short Version"

Set customShowToExpand = ActivePresentation.SlideShowSettings _
    .NamedSlideShows(customShowName)

slideToAddID = ActiveWindow.View.Slide.SlideID
customShowSlideIDs = customShowToExpand.SlideIDs
numSlides = UBound(customShowSlideIDs)

ReDim Preserve customShowSlideIDs(numSlides + 1)

customShowSlideIDs(numSlides + 1) = slideToAddID
customShowToExpand.Delete
ActivePresentation.SlideShowSettings.NamedSlideShows _
    .Add customShowName, customShowSlideIDs
```


## See also


#### Concepts


[NamedSlideShow Object](namedslideshow-object-powerpoint.md)

