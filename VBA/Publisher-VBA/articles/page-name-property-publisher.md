---
title: Page.Name Property (Publisher)
keywords: vbapb10.chm131098
f1_keywords:
- vbapb10.chm131098
ms.prod: publisher
api_name:
- Publisher.Page.Name
ms.assetid: cd81994d-506a-69ca-c7f6-472705b2ccd3
ms.date: 06/08/2017
---


# Page.Name Property (Publisher)

Returns or sets a  **String** value indicating the name of the specified object. Read/write.


## Syntax

 _expression_. **Name**

 _expression_A variable that represents a  **Page** object.


## Remarks

You can use an object's name in conjunction with the  **Item** method or **Item** property to return a reference to the object if the **Item** method or property for the collection that contains the object takes a **Variant** argument. For example, if the value of the **Name** property for a shape is Rectangle 2, then `.Shapes("Rectangle 2")` will return a reference to that shape.

The  **Name** property is the default property for the **BorderArt**,  **BorderArtFormat**, and  **Label** objects.


