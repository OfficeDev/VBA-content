---
title: Font.Name Property (PowerPoint)
keywords: vbapp10.chm575015
f1_keywords:
- vbapp10.chm575015
ms.prod: powerpoint
api_name:
- PowerPoint.Font.Name
ms.assetid: 6798b75b-7fb8-a046-1532-a8cc41b76af8
ms.date: 06/08/2017
---


# Font.Name Property (PowerPoint)

Returns or sets the name of the specified object. Read/write.


## Syntax

 _expression_. **Name**

 _expression_ A variable that represents a **Font** object.


### Return Value

String


## Remarks

You can use the object's name in conjunction with the  **Item** method to return a reference to the object if the **Item** method for the collection that contains the object takes a **Variant** argument. For example, if the value of the **Name** property for a shape is Rectangle 2, then `.Shapes("Rectangle 2")` will return a reference to that shape.


## See also


#### Concepts


[Font Object](font-object-powerpoint.md)

