---
title: Shape.Apply Method (PowerPoint)
keywords: vbapp10.chm547002
f1_keywords:
- vbapp10.chm547002
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.Apply
ms.assetid: 699a945f-656a-170a-e784-07b3004a858f
ms.date: 06/08/2017
---


# Shape.Apply Method (PowerPoint)

Applies to the specified shape formatting that's been copied by using the  **PickUp** method.


## Syntax

 _expression_. **Apply**

 _expression_ A variable that represents a **Shape** object.


## Example

This example copies the formatting of shape one on  `myDocument`, and then applies the copied formatting to shape two.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument

    .Shapes(1).PickUp

    .Shapes(2).Apply

End With
```


## See also


#### Concepts


[Shape Object](shape-object-powerpoint.md)

