---
title: ContainerProperties.InsertListMember Method (Visio)
keywords: vis_sdr.chm17662320
f1_keywords:
- vis_sdr.chm17662320
ms.prod: visio
api_name:
- Visio.ContainerProperties.InsertListMember
ms.assetid: be9c8bc6-7e2d-fb52-dd32-370a32d12744
ms.date: 06/08/2017
---


# ContainerProperties.InsertListMember Method (Visio)

Adds a shape or set of shapes to the list in the container.


## Syntax

 _expression_ . **InsertListMember**( **_ObjectToInsert_** , **_Position_** )

 _expression_ A variable that represents a **[ContainerProperties](containerproperties-object-visio.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ObjectToInsert_|Required| **[UNKNOWN]**|The object or selection to insert in the list. Can be a  **[Shape](shape-object-visio.md)** or **[Selection](selection-object-visio.md)** object.|
| _Position_|Required| **Long**|The insertion point in the list, which is one-based.|

### Return Value

 **Nothing**


## Remarks

If the container is not a list, Microsoft Visio returns an Invalid Source error.

If the  _ObjectToInsert_ parameter contains any non-top-level shapes or, if the list is locked, Visio returns an Invalid Parameter error. You cannot insert any of the following objects into a list:


- Another list.
    
- A callout.
    
- Routable connectors.
    
- Shapes whose PinX or PinY cells have the BOUND, GUARD, or SETATREF function applied.
    
- A member of another locked list.
    
Visio ignores shapes in  _ObjectToInsert_ that are already members of the list.

To insert before the first item in the list, pass 1 for the  _Position_ parameter. To insert at the end of the list, pass a _Position_ value equal to or greater than the value of the list count plus 1. If you pass an out-of-range position, Visio chooses the nearest valid position.

If  _ObjectToInsert_ does not match category requirements for lists, Visio returns an error.

Categories are user-defined strings that you can use to categorize shapes and, thereby, to restrict membership in a container. You can define categories in the User.msvShapeCategories cell in the ShapeSheet for a shape. You can define multiple categories for a shape by separating them with semicolons.


