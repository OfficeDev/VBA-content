---
title: Slides.FindBySlideID Method (PowerPoint)
keywords: vbapp10.chm530004
f1_keywords:
- vbapp10.chm530004
ms.prod: powerpoint
api_name:
- PowerPoint.Slides.FindBySlideID
ms.assetid: 49c5cb57-e132-0539-ecfd-25321ac7cc32
ms.date: 06/08/2017
---


# Slides.FindBySlideID Method (PowerPoint)

Returns a  **[Slide](slide-object-powerpoint.md)** object that represents the slide with the specified slide ID number. Each slide is automatically assigned a unique slide ID number when it is created. Use the **SlideID** property to return a slide's ID number.


## Syntax

 _expression_. **FindBySlideID**( **_SlideID_** )

 _expression_ A variable that represents a **Slides** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SlideID_|Required|**Long**|Specifies the ID number of the slide you want to return. Microsoft PowerPoint assigns this number when the slide is created.|

### Return Value

Slide


## Remarks

Unlike the  **SlideIndex** property, the **SlideID** property of a **Slide** object won't change when you add slides to the presentation or rearrange the slides in the presentation. Therefore, using the **FindBySlideID** method with the slide ID number can be a more reliable way to return a specific **Slide** object from a **[Slides](slides-object-powerpoint.md)** collection than using the **Item** method with the slide's index number.


## Example

This example demonstrates how to retrieve the unique ID number for a  **Slide** object and then use this number to return that **Slide** object from the **Slides** collection.


```vb
Set gslides = ActivePresentation.Slides

'Get slide ID
graphSlideID = gslides.Add(2, ppLayoutChart).SlideID

gslides.FindBySlideID(graphSlideID) _
    .SlideShowTransition.EntryEffect = _
    ppEffectCoverLeft      'Use ID to return specific slide
```


## See also


#### Concepts


[Slides Object](slides-object-powerpoint.md)

