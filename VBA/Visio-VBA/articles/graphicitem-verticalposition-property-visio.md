---
title: GraphicItem.VerticalPosition Property (Visio)
keywords: vis_sdr.chm16960445
f1_keywords:
- vis_sdr.chm16960445
ms.prod: visio
api_name:
- Visio.GraphicItem.VerticalPosition
ms.assetid: a756df97-851d-c6cf-b68f-b84e07da8628
ms.date: 06/08/2017
---


# GraphicItem.VerticalPosition Property (Visio)

Gets or sets the vertical position of the  **GraphicItem** object relative to the shape to which it is applied. Read/write.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **VerticalPosition**

 _expression_ An expression that returns a **GraphicItem** object.


### Return Value

VisGraphicPositionVertical


## Remarks

The default vertical callout position for graphic items in the  **GraphicItems** collection of a **Master** object of type **visTypeDataGraphic** , commonly called a data graphic master, is specified by the settings of the **Master.DataGraphicVerticalPosition** property. If **GraphicItem.UseDataGraphicPosition** is **True** , the graphic item is positioned according to the default setting. If **UseDataGraphicPosition** is **False** , its vertical position is determined by the settings of the **GraphicItem.VerticalPosition** property.

Note, however, that users can manually re-position a data graphic that has been applied to a shape by using the control handle of the data graphic. A position set in this manner takes precedence over the position specified by property settings.

The setting of the  **VerticalPosition** property corresponds to the setting of the **Vertical** list box under **Callout position** in the **New Text** (or **Edit Text**),  **New Icon Set** (or **Edit Icon Set**), or  **New Data Bar** (or **Edit Data Bar**) dialog box, depending on the type of the graphic item, in the Microsoft Visio user interface (UI). However, this is the case only if  **Use default position** is cleared.

The following possible values for the  **VerticalPosition** property are from the **VisGraphicPositionVertical** enumeration, which is declared in the Visio type library.



|**Constant**|**Value **|**Description**|
|:-----|:-----|:-----|
| **visGraphicBelow**|0|The top edge of the graphic item's alignment box is aligned with the bottom edge of the shape or container's alignment box.|
| **visGraphicBottomEdge**|1|The horizontal centerline of the graphic item's alignment box is aligned with the bottom edge of the shape or container's alignment box.|
| **visGraphicBottom**|2|The bottom edge of the graphic item's alignment box is aligned with the bottom edge of the shape or container's alignment box.|
| **visGraphicMiddle**|3|The horizontal centerline of the graphic item's alignment box is aligned with the horizontal centerline of the shape or container's alignment box.|
| **visGraphicTop**|4|The top edge of the graphic item's alignment box is aligned with the top edge of the shape or container's alignment box.|
| **visGraphicTopEdge**|5|The horizontal centerline of the graphic item's alignment box is aligned with the top edge of the shape or container's alignment box.|
| **visGraphicAbove**|6|The bottom edge of the graphic item's alignment box is aligned with the top edge of the shape or container's alignment box.|

 **Note**  Before you can set any property of a graphic item, you must use the  **[Master.Open](master-open-method-visio.md)** method to open for editing a copy of the data graphic master whose **GraphicItems** collection the graphic item belongs to. When you are finished setting properties, use the **Master.Close** method to commit changes.


