---
title: GraphicItem.HorizontalPosition Property (Visio)
keywords: vis_sdr.chm16960440
f1_keywords:
- vis_sdr.chm16960440
ms.prod: visio
api_name:
- Visio.GraphicItem.HorizontalPosition
ms.assetid: 268c461f-0290-3e3b-98f4-fa15bc902fa6
ms.date: 06/08/2017
---


# GraphicItem.HorizontalPosition Property (Visio)

Gets or sets the horizontal position of the  **GraphicItem** object relative to the shape to which it is applied. Read/write.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **HorizontalPosition**

 _expression_ An expression that returns a **GraphicItem** object.


### Return Value

 **VisGraphicPositionHorizontal**


## Remarks

The default horizontal callout position for graphic items in the  **GraphicItems** collection of a **Master** object of type **visTypeDataGraphic** , commonly called a data graphic master, is specified by the settings of the **Master.DataGraphicHorizontalPosition** property. If **GraphicItem.UseDataGraphicPosition** is **True** , the graphic item is positioned according to the default setting. If **UseDataGraphicPosition** is **False** , its horizontal position is determined by the settings of the **GraphicItem.HorizontalPosition** property.

Note, however, that users can manually re-position a data graphic that was applied to a shape by using the control handle of the data graphic. A position set in this manner takes precedence over the position specified by property settings.

The setting of the  **HorizontalPosition** property corresponds to the setting of the **Horizontal** list box under **Callout position** in the **New Text** (or **Edit Text**),  **New Icon Set** (or **Edit Icon Set**), or  **New Data Bar** (or **Edit Data Bar**) dialog box, depending on the type of the graphic item, in the Microsoft Visio user interface (UI). However, this is the case only if  **Use default position** is cleared.

The following possible values for the  **HorizontalPosition** property are from the **VisGraphicPositionHorizontal** enumeration, which is declared in the Visio type library.



|**Constant**|**Value **|**Description**|
|:-----|:-----|:-----|
| **visGraphicFarLeft**|0|The right edge of the graphic item's alignment box is aligned with the left edge of the shape or container's alignment box.|
| **visGraphicLeftEdge**|1|The vertical centerline of the graphic item's alignment box is aligned with the left edge of the shape or container's alignment box.|
| **visGraphicLeft**|2|The left edge of the graphic item's alignment box is aligned with the left edge of the shape or container's alignment box.|
| **visGraphicCenter**|3|The vertical centerline of the graphic item's alignment box is aligned with the vertical centerline of the shape or container's alignment box.|
| **visGraphicRight**|4|The right edge of the graphic item's alignment box is aligned with the right edge of the shape or container's alignment box.|
| **visGraphicRightEdge**|5|The vertical centerline of the graphic item's alignment box is aligned with the right edge of the shape or container's alignment box.|
| **visGraphicFarRight**|6|The left edge of the graphic item's alignment box is aligned with the right edge of the shape or container's alignment box.|

 **Note**  Before you can set any property of a graphic item, you must use the  **[Master.Open](master-open-method-visio.md)** method to open for editing a copy of the data graphic master whose **GraphicItems** collection the graphic item belongs to. When you are finished setting properties, use the **Master.Close** method to commit changes.


