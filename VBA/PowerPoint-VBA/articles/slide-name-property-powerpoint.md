---
title: Slide.Name Property (PowerPoint)
keywords: vbapp10.chm531008
f1_keywords:
- vbapp10.chm531008
ms.prod: powerpoint
api_name:
- PowerPoint.Slide.Name
ms.assetid: 11d6a295-02b6-3cf2-0e8b-42637e3b1f11
ms.date: 06/08/2017
---


# Slide.Name Property (PowerPoint)

When a slide is inserted into a presentation, Microsoft PowerPoint automatically assigns it a name in the form Slide _n_, where _n_ is an integer that represents the order in which the slide was created in the presentation. For example, the first slide inserted into a presentation is automatically named Slide1. If you copy a slide from one presentation to another, the slide loses the name it had in the first presentation and is automatically assigned a new name in the second presentation. A slide range must contain exactly one slide. Read/write **String**.


## Syntax

 _expression_. **Name**

 _expression_ A variable that represents a **Slide** object.


### Return Value

String


## Remarks

You can use the object's name in conjunction with the  **Item** method to return a reference to the object if the **Item** method for the collection that contains the object takes a **Variant** argument. For example, if the value of the **Name** property for a shape is Rectangle 2, then `.Shapes("Rectangle 2")` will return a reference to that shape.


## See also


#### Concepts


[Slide Object](slide-object-powerpoint.md)

