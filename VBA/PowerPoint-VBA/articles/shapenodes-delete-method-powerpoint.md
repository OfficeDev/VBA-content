---
title: ShapeNodes.Delete Method (PowerPoint)
keywords: vbapp10.chm560005
f1_keywords:
- vbapp10.chm560005
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeNodes.Delete
ms.assetid: a132067b-b8d7-0730-5dec-2df666eac209
ms.date: 06/08/2017
---


# ShapeNodes.Delete Method (PowerPoint)

Deletes a shape node.


## Syntax

 _expression_. **Delete**( **_Index_** )

 _expression_ A variable that represents a **ShapeNodes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Long**|Specifies the node to be deleted. |

## Remarks

The segment following the Index node is also deleted. If the node is a control point of a curve, the curve and all of its nodes are deleted.


## See also


#### Concepts


[ShapeNodes Object](shapenodes-object-powerpoint.md)

