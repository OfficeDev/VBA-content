---
title: Page.LayoutIncremental Method (Visio)
keywords: vis_sdr.chm10962140
f1_keywords:
- vis_sdr.chm10962140
ms.prod: visio
api_name:
- Visio.Page.LayoutIncremental
ms.assetid: db112261-120d-e2e8-18f0-91b1bba0a3a4
ms.date: 06/08/2017
---


# Page.LayoutIncremental Method (Visio)

Makes small adjustments to the position of shapes on the drawing page to better align the shapes or to space them evenly from other shapes.


## Syntax

 _expression_ . **LayoutIncremental**( **_AlignOrSpace_** , **_AlignHorizontal_** , **_AlignVertical_** , **_SpaceHorizontal_** , **_SpaceVertical_** , **_UnitsNameOrCode_** )

 _expression_ A variable that represents a **[Page](page-object-visio.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _AlignOrSpace_|Required| **[VisLayoutIncrementalType](vislayoutincrementaltype-enumeration-visio.md)**|The type of incremental layout action to perform--alignment, spacing, or both. See Remarks for possible values.|
| _AlignHorizontal_|Required| **[VisLayoutHorzAlignType](vislayouthorzaligntype-enumeration-visio.md)**|Indicates how Microsoft Visio aligns shapes horizontally when it aligns incrementally. See Remarks for possible values.|
| _AlignVertical_|Required| **[VisLayoutVertAlignType](vislayoutvertaligntype-enumeration-visio.md)**|Indicates how Visio aligns shapes vertically when it aligns incrementally (except if layout style is circular). See Remarks for possible values.|
| _SpaceHorizontal_|Required| **Double**|The edge-to-edge horizontal spacing. Must be greater than or equal to zero.|
| _SpaceVertical_|Required| **Double**|The edge-to-edge vertical spacing (except if layout style is circular). Must be greater than or equal to zero.|
| _UnitsNameOrCode_|Required| **[VisUnitCodes](visunitcodes-enumeration-visio.md)**|The units for the spacing values.|

### Return Value

 **Nothing**


## Remarks

The  _AlignOrSpace_ parameter must be one or the combination of both (3) of the following **VisLayoutIncrementalType** constants.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visLayoutIncrAlign**|1|Align shapes.|
| **visLayoutIncrSpace**|2|Space shapes evenly.|
The  _AlignHorizontal_ parameter must be one of the following **VisLayoutHorzAlignType** constants.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visLayoutHorzAlignNone**|0|Do not align horizontally.|
| **visLayoutHorzAlignDefault**|1|Visio chooses how to align horizontally.|
| **visLayoutHorzAlignLeft**|2|Align the left edges of the shapes.|
| **visLayoutHorzAlignCenter**|3|Align the centers of the shapes.|
| **visLayoutHorzAlignRight**|4|Align the right edges of the shapes.|
The  _AlignVertical_ parameter must be one of the following **VisLayoutVertAlignType** constants.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visLayoutVertAlignNone**|0|Do not align vertically.|
| **visLayoutVertAlignDefault**|1|Visio chooses how to align vertically.|
| **visLayoutVertAlignTop**|2|Align the top edges of the shapes.|
| **visLayoutVertAlignMiddle**|3|Align the middles of the shapes.|
| **visLayoutVertAlignBottom**|4|Align the bottom edges of the shapes.|
If  _AlignOrSpace_ is **visLayoutIncrAlign** , either _AlignHorizontal_ must be a value other than **visLayoutHorzAlignNone** , or _AlignVertical_ must be a value other than **visLayoutVertAlignNone** .

If  _AlignOrSpace_ is **visLayoutIncrSpace** , both _AlignHorizontal_ and _AlignVertical_ must be greater than zero.

If  _AlignOrSpace_ is a combination of **visLayoutIncrAlign** and **visLayoutIncrSpace** , both of these conditions must be true.

If the page layout style is circular, Visio uses only the  _AlignHorizontal_ value to determine whether to align, and only the _SpaceHorizontal_ value to determine whether to space, ignoring the _AlignVertical_ and _SpaceVertical_ values, respectively. In this case, if you pass anything other than **visLayoutHorzAlignNone** for _AlignHorizontal_ , Visio performs the alignment. Similarly, if you pass any value greater than zero for _SpaceHorizontal_ , Visio performs the spacing.


