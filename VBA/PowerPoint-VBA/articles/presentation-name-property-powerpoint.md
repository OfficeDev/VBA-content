---
title: Presentation.Name Property (PowerPoint)
keywords: vbapp10.chm583025
f1_keywords:
- vbapp10.chm583025
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.Name
ms.assetid: a93a6d21-e3e7-0d7d-ae73-34f9511445de
ms.date: 06/08/2017
---


# Presentation.Name Property (PowerPoint)

The name of the presentation includes the file name extension (for file types that are registered) but doesn't include its path. You cannot use this property to set the name. Use the  **[SaveAs](presentation-saveas-method-powerpoint.md)** method to save the presentation under a different name if you need to change the name. Read-only.


## Syntax

 _expression_. **Name**

 _expression_ A variable that represents a **Presentation** object.


### Return Value

String


## Remarks

You can use the object's name in conjunction with the  **Item** method to return a reference to the object if the **Item** method for the collection that contains the object takes a **Variant** argument. For example, if the value of the **Name** property for a shape is Rectangle 2, then `.Shapes("Rectangle 2")` will return a reference to that shape.


## See also


#### Concepts


[Presentation Object](presentation-object-powerpoint.md)

