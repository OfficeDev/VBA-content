---
title: GraphicItem.UseDataGraphicPosition Property (Visio)
keywords: vis_sdr.chm16960450
f1_keywords:
- vis_sdr.chm16960450
ms.prod: visio
api_name:
- Visio.GraphicItem.UseDataGraphicPosition
ms.assetid: d463eefb-8103-3701-fd8a-604c65f74713
ms.date: 06/08/2017
---


# GraphicItem.UseDataGraphicPosition Property (Visio)

Gets or sets whether to use the current default callout position for graphic items of the data graphic master to whose  **GraphicItems** collection the graphic item belongs. Read/write.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **UseDataGraphicPosition**

 _expression_ An expression that returns a **GraphicItem** object.


### Return Value

Boolean


## Remarks

The default callout position for graphic items in the  **GraphicItems** collection of a **Master** object of type **visTypeDataGraphic** , commonly called a data graphic master, is specified by the settings of the **Master.DataGraphicVerticalPosition** and **Master.DataGraphicHorizontalPosition** properties. If **GraphicItem.UseDataGraphicPosition** is **True** , the graphic item is positioned according to the default setting. If **UseDataGraphicPosition** is **False** , its position is determined by the settings of the **GraphicItem.VerticalPosition** and **GraphicItem.HorizontalPosition** properties.

In addition, if the  **HorizontalPosition** and **VerticalPosition** property values of a graphic item are equal to the **DataGraphicHorizontalPosition** and **DataGraphicVerticalPosition** property values, the value of the **UseDataGraphicPosition** property for that graphic item is automatically set to **True** .

Note, however, that users can manually re-position a data graphic that has been applied to a shape by using the control handle of the data graphic. A position set in this manner takes precedence over the position specified by property settings.


 **Note**  Before you can set any property of a graphic item, you must use the  **[Master.Open](master-open-method-visio.md)** method to open for editing a copy of the data graphic master whose **GraphicItems** collection the graphic item belongs to. When you are finished setting properties, use the **Master.Close** method to commit changes.

The setting of the  **UseDataGraphicPosition** property corresponds to the setting of the **Use default position** box under **Callout position** in the **New Text** (or **Edit Text**),  **New Icon Set** (or **Edit Icon Set**), or  **New Data Bar** (or **Edit Data Bar**) dialog box, depending on the type of the graphic item, in the Microsoft Visio user interface.


