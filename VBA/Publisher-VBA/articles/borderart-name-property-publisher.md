---
title: BorderArt.Name Property (Publisher)
keywords: vbapb10.chm7667715
f1_keywords:
- vbapb10.chm7667715
ms.prod: publisher
api_name:
- Publisher.BorderArt.Name
ms.assetid: 71e68f08-6bf1-7f4d-5da2-9f4ce4f63021
ms.date: 06/08/2017
---


# BorderArt.Name Property (Publisher)

Returns a  **String** value indicating the name of the specified object. Read-only.


## Syntax

 _expression_. **Name**

 _expression_A variable that represents a  **BorderArt** object.


## Remarks

You can use an object's name in conjunction with the  **Item** method or **Item** property to return a reference to the object if the **Item** method or property for the collection that contains the object takes a **Variant** argument. For example, if the value of the **Name** property for a shape is Rectangle 2, then `.Shapes("Rectangle 2")` will return a reference to that shape.

The  **Name** property is the default property for the **BorderArt**,  **BorderArtFormat**, and  **Label** objects.


## See also


#### Concepts


 [BorderArt Object](borderart-object-publisher.md)

