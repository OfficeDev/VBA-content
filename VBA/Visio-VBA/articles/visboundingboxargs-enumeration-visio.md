---
title: VisBoundingBoxArgs Enumeration (Visio)
keywords: vis_sdr.chm70060
f1_keywords:
- vis_sdr.chm70060
ms.prod: visio
ms.assetid: 04523cbd-758f-757d-daac-30ca4676e6c2
ms.date: 06/08/2017
---


# VisBoundingBoxArgs Enumeration (Visio)

Flags to be passed to the  **BoundingBox** method of various objects.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visBBoxDrawingCoords**|&;H2000|Return numbers in the drawing coordinate system of the page or master whose shapes are being considered. By default, the returned numbers are drawing units in the local coordinate system of the parent of the considered shapes.|
| **visBBoxExtents**|&;H4|Return a rectangle that is the smallest rectangle parallel to the local coordinate system of the shape's parent that encloses the paths stroked by the shape's geometry.This rectangle may be larger or smaller than the shape's upright width-height box. The extents box determined for a shape of type  **visTypeForeignObject** equals that shape's upright width-height box.|
| **visBBoxIgnoreVisible**|&;H20|Ignore visible geometry.|
| **visBBoxIncludeDataGraphics**|&;H10000|Include data-graphic callout shapes (and their sub-shapes) that are applied to the shape, or the shapes in a master, page, or selection. Off by default.|
| **visBBoxIncludeGuides**|&;H1000|Include extents for shapes of type  **visTypeguide** . By default, the extents of shapes of type **visTypeGuide** are ignored.If you request guide extents, only the positions of vertical guides and the positions of horizontal guides contribute to the rectangle that is returned. If any vertical guides are reported on, an infinite extent is returned. If any horizontal guides are reported on, an infinite extent is returned. If any rotated guides are reported on, infinite and extents are returned.|
| **visBBoxIncludeHidden**|&;H10|Include hidden geometry.|
| **visBBoxNoNonPrint**|&;H4000|Ignore the extents of shapes that are non-printing. A shape is non-printing if the value of its NonPrinting cell is non-zero or it belongs only to non-printing layers.|
| **visBBoxUprightText**|&;H2|Return a rectangle that is the smallest rectangle parallel to the local coordinate system of the shape's parent that encloses the shape's text.|
| **visBBoxUprightWH**|&;H1|Return a rectangle that is the smallest rectangle parallel to the local coordinate system of the shape's parent that encloses the shape's width-height box.If the shape is not rotated, its upright width-height box and its width-height box are the same. Paths in the shape's geometry need not and often do not lie entirely within the shape's width-height box.|

