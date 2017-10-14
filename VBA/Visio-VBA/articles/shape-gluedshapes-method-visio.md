---
title: Shape.GluedShapes Method (Visio)
keywords: vis_sdr.chm11262245
f1_keywords:
- vis_sdr.chm11262245
ms.prod: visio
api_name:
- Visio.Shape.GluedShapes
ms.assetid: 0c9c551d-ce28-f7c6-4656-8120962e8d34
ms.date: 06/08/2017
---


# Shape.GluedShapes Method (Visio)

Returns an array that contains the identifiers of the shapes that are glued to a shape.


## Syntax

 _expression_ . **GluedShapes**( **_Flags_** , **_CategoryFilter_** , **_pOtherConnectedShape_** )

 _expression_ A variable that represents a **[Shape](shape-object-visio.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Flags_|Required| **[VisGluedShapesFlags](visgluedshapesflags-enumeration-visio.md)**|The dimensionality and directionality of the connection points of the shapes to return.|
| _CategoryFilter_|Required| **String**|The category of shapes to return. See Remarks for more information|
| _pOtherConnectedShape_|Optional| **Shape**|Additional shape to which returned shapes must also be glued. |

### Return Value

 **Long()**


## Remarks

 _Flags_ must be one of the following **VisGluedShapesFlags** constants.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visGluedShapesAll1D**|0|Return all 1-D shapes that are glued to this shape.|
| **visGluedShapesIncoming1D**|1|Return 1-D shapes whose end points are glued to this shape.|
| **visGluedShapesOutgoing1D**|2|Return 1-D shapes whose begin points are glued to this shape.|
| **visGluedShapesAll2D**|3|Return all 2-D shapes that are glued to this shape and all 2-D shapes to which this shape is glued. |
| **visGluedShapesIncoming2D**|4|If the source object is a 1-D shape, return the 2-D shape to which the begin point is glued. If the source object is a 2-D shape, return the 2-D shapes that are glued to this shape.|
| **visGluedShapesOutgoing2D**|5|If the source object is a 1-D shape, return the 2-D shape to which the end point is glued. If the source object is a 2-D shape, return the 2-D shapes to which this shape is glued.|
Categories are user-defined strings that you can use to categorize shapes and thereby to restrict membership in a container. You can define categories in the User.msvShapeCategories cell in the ShapeSheet for a shape. You can define multiple categories for a shape by separating those categories with semi-colons.

Connection points with dual directionality (both inward and outward) are identified as incoming or outgoing based on the way that they are used in a particular connection.

The  **GluedShapes** method fails if the source object is a part of a master or a guide. Guides are excluded from any list of shapes returned.

If you specify an invalid shape for  _pOtherConnectedShape_ , Microsoft Visio returns an Invalid Parameter error.

 **GluedShapes** returns an empty array if there are no qualifying shapes to return.


