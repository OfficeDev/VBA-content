---
title: SoundEffect.Name Property (PowerPoint)
keywords: vbapp10.chm540003
f1_keywords:
- vbapp10.chm540003
ms.prod: powerpoint
api_name:
- PowerPoint.SoundEffect.Name
ms.assetid: f587126e-094a-0360-b696-fbdb7c0a4019
ms.date: 06/08/2017
---


# SoundEffect.Name Property (PowerPoint)

Returns or sets the name of the specified object. Read/write.


## Syntax

 _expression_. **Name**

 _expression_ A variable that represents a **SoundEffect** object.


### Return Value

String


## Remarks

You can use the object's name in conjunction with the  **Item** method to return a reference to the object if the **Item** method for the collection that contains the object takes a **Variant** argument. For example, if the value of the **Name** property for a shape is Rectangle 2, then `.Shapes("Rectangle 2")` will return a reference to that shape.


## See also


#### Concepts


[SoundEffect Object](soundeffect-object-powerpoint.md)

