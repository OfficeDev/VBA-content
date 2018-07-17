---
title: Application.Name Property (PowerPoint)
keywords: vbapp10.chm502009
f1_keywords:
- vbapp10.chm502009
ms.prod: powerpoint
api_name:
- PowerPoint.Application.Name
ms.assetid: c7a59327-774a-8c55-17b4-053ae76bd623
ms.date: 06/08/2017
---


# Application.Name Property (PowerPoint)

Returns the string "Microsoft PowerPoint." Read-only.


## Syntax

 _expression_. **Name**

 _expression_ A variable that represents an **Application** object.


### Return Value

String


## Remarks

You can use the object's name in conjunction with the  **Item** method to return a reference to the object if the **Item** method for the collection that contains the object takes a **Variant** argument. For example, if the value of the **Name** property for a shape is Rectangle 2, then `.Shapes("Rectangle 2")` will return a reference to that shape.


## See also


#### Concepts


[Application Object](application-object-powerpoint.md)

