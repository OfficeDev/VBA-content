---
title: NamedSlideShow.Name Property (PowerPoint)
keywords: vbapp10.chm516003
f1_keywords:
- vbapp10.chm516003
ms.prod: powerpoint
api_name:
- PowerPoint.NamedSlideShow.Name
ms.assetid: fda5a218-764e-3792-809c-14d9e9da1ce2
ms.date: 06/08/2017
---


# NamedSlideShow.Name Property (PowerPoint)

You cannot use this property to set the name for a custom slide show. Use the  **[Add](namedslideshows-add-method-powerpoint.md)** method to redefine a custom slide show under a new name. Read-only.


## Syntax

 _expression_. **Name**

 _expression_ A variable that represents a **NamedSlideShow** object.


### Return Value

String


## Remarks

You can use the object's name in conjunction with the  **Item** method to return a reference to the object if the **Item** method for the collection that contains the object takes a **Variant** argument. For example, if the value of the **Name** property for a shape is Rectangle 2, then `.Shapes("Rectangle 2")` will return a reference to that shape.


## See also


#### Concepts


[NamedSlideShow Object](namedslideshow-object-powerpoint.md)

