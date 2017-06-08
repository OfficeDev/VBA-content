---
title: ShapeRange.Align Method (Publisher)
keywords: vbapb10.chm2294016
f1_keywords:
- vbapb10.chm2294016
ms.prod: publisher
api_name:
- Publisher.ShapeRange.Align
ms.assetid: ef522d47-3fc7-cfca-5b9a-44ff020f8b31
ms.date: 06/08/2017
---


# ShapeRange.Align Method (Publisher)

Aligns all the shapes in the specified  **ShapeRange** object.


## Syntax

 _expression_. **Align**( **_AlignCmd_**,  **_RelativeTo_**)

 _expression_A variable that represents a  **ShapeRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|AlignCmd|Required| **MsoAlignCmd**|Specifies how the shapes are to be aligned.|
|RelativeTo|Required| **MsoTriState**|Specifies whether shapes are aligned relative to the page or to one another.|

## Remarks

If the RelativeTo parameter is  **msoFalse** and the shape range contains only one shape, an error occurs.

The AlignCmd parameter can be one of the  **MsoAlignCmd** constants declared in the Microsoft Office type library.



|**Constant**|**Description**|
|:-----|:-----|
| **msoAlignBottoms**|Aligns shapes along their bottom edges. If  _RelativeTo_ is **msoFalse**, the bottommost shape determines the line against which the other shapes are aligned.|
| **msoAlignCenters**|Aligns shapes on a vertical line through their centers. If  _RelativeTo_ is **msoFalse**, shapes are aligned on a line halfway between the left- and rightmost shapes.|
| **msoAlignLefts**|Aligns shapes along their left edges. If  _RelativeTo_ is **msoFalse**, the leftmost shape determines the line against which the other shapes are aligned.|
| **msoAlignMiddles**|Aligns shapes on a horizontal line through their centers. If  _RelativeTo_ is **msoFalse**, shapes are aligned on a line halfway between the top- and bottommost shapes.|
| **msoAlignRights**| **msoAlignRights** Aligns shapes along their right edges. If _RelativeTo_ is **msoFalse**, the rightmost shape determines the line against which the other shapes are aligned.|
| **msoAlignTops**| Aligns shapes along their top edges. If _RelativeTo_ is **msoFalse**, the topmost shape determines the line against which the other shapes are aligned.|
The RelativeTo parameter can be one of the  **MsoTriState** constants declared in the Microsoft Office type library and shown in the following table.



|**Constant**|**Description**|
|:-----|:-----|
| **msoFalse**|Aligns shapes relative to one another.|
| **msoTrue**|Aligns shapes relative to the page.|

## Example

The following example aligns all the shapes on the first page of the active publication on a vertical line through their centers.


```vb
ActiveDocument.Pages(1).Shapes.Range.Align _ 
 AlignCmd:=msoAlignCenters, _ 
 RelativeTo:=msoTrue 

```


