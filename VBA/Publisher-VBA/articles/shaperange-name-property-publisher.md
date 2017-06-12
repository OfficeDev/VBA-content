---
title: ShapeRange.Name Property (Publisher)
keywords: vbapb10.chm2293828
f1_keywords:
- vbapb10.chm2293828
ms.prod: publisher
api_name:
- Publisher.ShapeRange.Name
ms.assetid: 517eca4b-fa8c-0f6a-2829-75704bb4c899
ms.date: 06/08/2017
---


# ShapeRange.Name Property (Publisher)

Returns or sets a  **String** value indicating the name of the specified object. Read/write.


## Syntax

 _expression_. **Name**

 _expression_A variable that represents a  **ShapeRange** object.


## Remarks

You can use an object's name in conjunction with the  **Item** method or **Item** property to return a reference to the object if the **Item** method or property for the collection that contains the object takes a **Variant** argument. For example, if the value of the **Name** property for a shape is Rectangle 2, then `.Shapes("Rectangle 2")` will return a reference to that shape.

The  **Name** property is the default property for the **BorderArt**,  **BorderArtFormat**, and  **Label** objects.


