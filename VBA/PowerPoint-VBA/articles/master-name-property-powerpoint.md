---
title: Master.Name Property (PowerPoint)
keywords: vbapp10.chm533007
f1_keywords:
- vbapp10.chm533007
ms.prod: powerpoint
api_name:
- PowerPoint.Master.Name
ms.assetid: 1c751814-61fe-c246-d516-0d43b7757248
ms.date: 06/08/2017
---


# Master.Name Property (PowerPoint)

Returns or sets the name of the specified object. Read/write.


## Syntax

 _expression_. **Name**

 _expression_ A variable that represents a **Master** object.


### Return Value

String


## Remarks

You can use the object's name in conjunction with the  **Item** method to return a reference to the object if the **Item** method for the collection that contains the object takes a **Variant** argument. For example, if the value of the **Name** property for a shape is Rectangle 2, then `.Shapes("Rectangle 2")` will return a reference to that shape.


## See also


#### Concepts


[Master Object](master-object-powerpoint.md)

