---
title: Page.DropIntoList Method (Visio)
keywords: vis_sdr.chm10962180
f1_keywords:
- vis_sdr.chm10962180
ms.prod: visio
api_name:
- Visio.Page.DropIntoList
ms.assetid: fcefca11-d64b-9f95-a00e-bf9968d26267
ms.date: 06/08/2017
---


# Page.DropIntoList Method (Visio)

Drops the specified object into the specified list at the specified position. Returns the newly dropped shape.


## Syntax

 _expression_ . **DropIntoList**( **_ObjectToDrop_** , **_TargetList_** , **_lPosition_** )

 _expression_ An expression that returns a **[Page](page-object-visio.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ObjectToDrop_|Required| **IUnknown**|The source of the shape to drop into the list. Can be a  **[Master](master-object-visio.md)** , **[Selection](selection-object-visio.md)** , **[Shape](shape-object-visio.md)** , or **IDataObject** object. See Remarks for more information.|
| _TargetList_|Required| **Shape**|The list into which to drop  _ObjectToDrop_. |
| _lPosition_|Required| **Long**|The position in the 1-based list to add the shape.|

### Return Value

 **Shape**


## Remarks

If  _ObjectToDrop_ is a **Selection** object, the selection can contain only a single shape.

If  _ObjectToDrop_ is an **IDataObject** , it must be associated with a local Microsoft Visio object that is in the same instance as the page on which it is being dropped.

Visio returns an Invalid Target error if  _ObjectToDrop_ does not match the category requirements of the list or the container. Shapes can be assigned categories, and containers can have required and excluded categories.

Categories are user-defined strings that you can use to categorize shapes and, thereby, to restrict membership in a container. You can define categories in the User.msvShapeCategories cell in the ShapeSheet for a shape. You can define multiple categories for a shape by separating the categories with semicolons. 

If  _ObjectToDrop_ is not a Microsoft Visio object, or if it does not contain top-level shapes on the page, Microsoft Visio returns an Invalid Parameter error.

If the  **[ContainerProperties.LockMembership](containerproperties-lockmembership-property-visio.md)** property of the list is **True** , Visio returns a Disabled error.


## Example

The following Visual Basic for Applications (VBA) example shows how to use the  **DropIntoList** method to add a new shape to an existing list on the active page, in the first position in the list.


```vb
Application.ActivePage.DropIntoList vsoMaster, vsoListShape, 1
```


