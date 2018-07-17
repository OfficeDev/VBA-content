---
title: Slides Object (PowerPoint)
keywords: vbapp10.chm530000
f1_keywords:
- vbapp10.chm530000
ms.prod: powerpoint
api_name:
- PowerPoint.Slides
ms.assetid: ba7f514c-8f6d-d5ef-333f-c1da0f2ab767
ms.date: 06/08/2017
---


# Slides Object (PowerPoint)

A collection of all the  **[Slide](slide-object-powerpoint.md)** objects in the specified presentation.


## Remarks

If your Visual Studio solution includes the  **Microsoft.Office.Interop.PowerPoint** reference, this collection maps to the following types:


-  **Microsoft.Office.Interop.PowerPoint.Slides.GetEnumerator** (to enumerate the **Slide** objects.)
    
The following examples describe how to:


- Create a slide and add it to the collection
    
- Return a single slide that you specify by name, index number, or slide ID number
    
- Return a subset of the slides in the presentation
    
- Apply a property or method to all the slides in the presentation at the same time
    

## Example

Use the [Slides](presentation-slides-property-powerpoint.md) property to return a **Slides** collection. Use the[Add](presentations-add-method-powerpoint.md) method to create a new slide and add it to the collection. The following example adds a new slide to the active presentation.


```vb
ActivePresentation.Slides.Add 2, ppLayoutBlank
```

Use  **Slides** (index), where index is the slide name or index number, or use the **Slides.FindBySlideID** (index), where index is the slide ID number, to return a single **Slide** object. The following example sets the layout for slide one in the active presentation.




```vb
ActivePresentation.Slides(1).Layout = ppLayoutTitle
```

The following example sets the layout for the slide named "Big Chart" in the active presentation. Note that slides are assigned automatically generated names of the form Sliden (where n is an integer) when they're created. To assign a more meaningful name to a slide, use the [Name](slide-name-property-powerpoint.md) property.




```vb
ActivePresentation.Slides("Big Chart").Layout = ppLayoutTitle
```

Use  **Slides.Range** (index), where index is the slide index number or name or an array of slide index numbers or an array of slide names, to return a **[SlideRange](sliderange-object-powerpoint.md)** object that represents a subset of the **Slides** collection. The following example sets the background fill for slides one and three in the active presentation.




```vb
With ActivePresentation.Slides.Range(Array(1, 3)) 
    .FollowMasterBackground = False 
    .Background.Fill.PresetGradient msoGradientHorizontal, _ 
        1, msoGradientLateSunset 
End With
```

If you want to do something to all the slides in your presentation at the same time (such as delete all of them or set a property for all of them), use  **Slides.Range** with no argument to construct a **SlideRange** collection that contains all the slides in the **Slides** collection, and then apply the appropriate property or method to the **SlideRange** collection. The following example sets the background fill for all the slides in the active presentation




```vb
With ActivePresentation.Slides.Range 
    .FollowMasterBackground = False 
    .Background.Fill.PresetGradient msoGradientHorizontal, _ 
        1, msoGradientLateSunset 
End With
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

