---
title: Tags Object (PowerPoint)
keywords: vbapp10.chm611000
f1_keywords:
- vbapp10.chm611000
ms.prod: powerpoint
api_name:
- PowerPoint.Tags
ms.assetid: 75ecbd43-0aa7-d49d-f1f5-c6c21d8babee
ms.date: 06/08/2017
---


# Tags Object (PowerPoint)

Represents a tag or a custom property that you can create for a shape, slide, or presentation. 


## Remarks

Each  **Tags** object contains the name of a custom property and a value for that property.

Create tags when you want to be able to selectively work with specific members of a collection, based on an attribute that isn't already represented by a built-in property. For example, if you want to be able to categorize slides in a presentation based on what region of the country/region they apply to, you could create a Region tag and assign a Region value to each slide in the presentation. You could then selectively perform an operation on some of the slides, based on the values of their Region tags, such as hiding all the slides with the Region value "East."


## Example

Use the [Add](tags-add-method-powerpoint.md) method to add a tag to an object. The following example adds a tag with the name "Region" and with the value "East" to slide one in the active presentation.


```vb
ActivePresentation.Slides(1).Tags.Add "Region", "East"
```

Use  **Tags** (index), where index is the name of a tag, to return a the tag value. The following example tests the value of the Region tag for all slides in the active presentation and hides any slides that don't pertain to the East Coast (denoted by the value "East").




```vb
For Each s In ActivePresentation.Slides

    If s.Tags("region") <> "east" Then

        s.SlideShowTransition.Hidden = True

    End If

Next
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

