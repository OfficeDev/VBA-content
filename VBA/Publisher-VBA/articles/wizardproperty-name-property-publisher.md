---
title: WizardProperty.Name Property (Publisher)
keywords: vbapb10.chm1572864
f1_keywords:
- vbapb10.chm1572864
ms.prod: publisher
api_name:
- Publisher.WizardProperty.Name
ms.assetid: d66dd4be-9f47-baed-b4aa-6c8cbf293505
ms.date: 06/08/2017
---


# WizardProperty.Name Property (Publisher)

Returns a  **String** value indicating the name of the specified object. Read-only.


## Syntax

 _expression_. **Name**

 _expression_A variable that represents a  **WizardProperty** object.


## Remarks

You can use an object's name in conjunction with the  **Item** method or **Item** property to return a reference to the object if the **Item** method or property for the collection that contains the object takes a **Variant** argument. For example, if the value of the **Name** property for a shape is Rectangle 2, then `.Shapes("Rectangle 2")` will return a reference to that shape.

The  **Name** property is the default property for the **BorderArt**,  **BorderArtFormat**, and  **Label** objects.


