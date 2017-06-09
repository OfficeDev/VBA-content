---
title: MsoMergeCmd Enumeration (Office)
ms.assetid: 41a8359d-9a48-f847-488c-b842efe15e28
ms.date: 06/08/2017
ms.prod: office
---


# MsoMergeCmd Enumeration (Office)

Specifies the output of a merge shapes operation.


## Members



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
|**msoMergeCombine**|**2**|Creates a new shape from selected shapes. If the selected shapes overlap, the area where they overlap is cut out, or discarded. |
|**msoMergeFragment**|**5**|Breaks a shape into smaller parts or create new shapes from intersecting lines or from shapes that overlap.|
|**msoMergeIntersect**|**3**|Forms a new closed shape from the area where selected shapes overlap, eliminating non-overlapping areas.|
|**msoMergeSubtract**|**4**|Creates a new shape by subtracting from the primary selection the areas where subsequent selections overlap. |
|**msoMergeUnion**|**1**|Creates a new shape from the perimeter of two or more overlapping shapes. The new shape is a set of all the points from the original shapes.|
|**msoMergeCombine**|**2**||
|**msoMergeFragment**|**5**||
|**msoMergeIntersect**|**3**||
|**msoMergeSubtract**|**4**||
|**msoMergeUnion**|**1**||
|Name|Value|Description|

