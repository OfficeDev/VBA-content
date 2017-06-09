---
title: Master.DataGraphicHorizontalPosition Property (Visio)
keywords: vis_sdr.chm10760125
f1_keywords:
- vis_sdr.chm10760125
ms.prod: visio
api_name:
- Visio.Master.DataGraphicHorizontalPosition
ms.assetid: d9c98a41-ffc0-152e-2150-0915bd38bcac
ms.date: 06/08/2017
---


# Master.DataGraphicHorizontalPosition Property (Visio)

Gets or sets the default horizontal callout position for members of the  **GraphicItems** collection of the **Master** object of type **visTypeDataGraphic** . Read/write.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **DataGraphicHorizontalPosition**

 _expression_ An expression that returns a **Master** object.


### Return Value

VisGraphicPositionHorizontal


## Remarks

The default horizontal callout position for graphic items in the  **GraphicItems** collection of a **Master** object of type **visTypeDataGraphic** , commonly called a data graphic master, is specified by the settings of the **Master.DataGraphicHorizontalPosition** property. If **GraphicItem.UseDataGraphicPosition** is **True** , the graphic item is positioned according to the default setting. If **UseDataGraphicPosition** is **False** , its horizontal position is determined by the settings of the **GraphicItem.HorizontalPosition** property.

Note, however, that users can manually re-position a data graphic that was applied to a shape by using the control handle of the data graphic. A position set in this manner takes precedence over the position specified by property settings.

The setting of the  **DataGraphicHorizontalPosition** property corresponds to the setting of the **Horizontal** box under **Default position** in the **New Data Graphic** or **Edit Data Graphic** dialog box in the Microsoft Visio user interface (UI).

The following possible values for the  **DataGraphicHorizontalPosition** property are from the **VisGraphicPositionHorizontal** enumeration, which is declared in the Visio type library.



|**Constant**|**Value **|**Description**|
|:-----|:-----|:-----|
| **visGraphicFarLeft**|0|The right edge of the graphic item's alignment box is aligned with the left edge of the shape's alignment box.|
| **visGraphicLeftEdge**|1|The vertical centerline of the graphic item's alignment box is aligned with the left edge of the shape's alignment box.|
| **visGraphicLeft**|2|The left edge of the graphic item's alignment box is aligned with the left edge of the shape's alignment box.|
| **visGraphicCenter**|3|The vertical centerline of the graphic item's alignment box is aligned with the vertical centerline of the shape's alignment box.|
| **visGraphicRight**|4|The right edge of the graphic item's alignment box is aligned with the right edge of the shape's alignment box.|
| **visGraphicRightEdge**|5|The vertical centerline of the graphic item's alignment box is aligned with the right edge of the shape's alignment box.|
| **visGraphicFarRight**|6|The left edge of the graphic item's alignment box is aligned with the right edge of the shape's alignment box.|

 **Note**  Before you can set any property of a graphic item, you must use the  **[Master.Open](master-open-method-visio.md)** method to open for editing a copy of the data graphic master whose **GraphicItems** collection the graphic item belongs to. When you are finished setting properties, use the **Master.Close** method to commit changes.


